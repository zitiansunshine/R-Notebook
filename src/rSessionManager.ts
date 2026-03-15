// =============================================================================
// rSessionManager.ts — manages a persistent R subprocess per document
// =============================================================================

import * as cp    from 'child_process';
import * as crypto from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path  from 'path';
import * as readline from 'readline';
import { EventEmitter } from 'events';
import {
  KernelRequest, KernelResponse,
  ExecResult, DfDataResult, VarsResult,
} from './kernelProtocol';

interface PendingRequest {
  resolve: (r: KernelResponse) => void;
  reject:  (e: Error) => void;
  timeout: ReturnType<typeof setTimeout> | undefined;
}

const IDLE_CHECKPOINT_DELAY_MS = 10_000;

function sanitizeExecTimeoutMs(timeoutMs?: number): number {
  if (typeof timeoutMs === 'number' && timeoutMs === 0) return 0;
  return typeof timeoutMs === 'number' && Number.isFinite(timeoutMs) && timeoutMs > 0
    ? timeoutMs
    : 0;
}

export class RSession extends EventEmitter {
  private proc:    cp.ChildProcess | null = null;
  private rl:      readline.Interface | null = null;
  private pending: Map<string, PendingRequest> = new Map();
  private started  = false;
  private procId   = 0;
  private readonly kernelPath: string;
  private requestedBin: string;
  private activeBin?: string;
  private expectedExitProcId: number | null = null;
  private activeExecKey: string | null = null;
  private interruptingExecKey: string | null = null;
  private interruptRetryTimer: ReturnType<typeof setInterval> | null = null;
  private interruptEscalationTimer: ReturnType<typeof setTimeout> | null = null;
  private interruptAttempts = 0;
  private lastInterruptAt = 0;
  private lastVars: VarsResult = { type: 'vars_result', vars: [] };
  private readonly checkpointPath: string;
  private workspaceDirty = false;
  private hasCheckpoint = false;
  private checkpointingDisabled = false;
  private recoveryPromise: Promise<void> | null = null;
  private checkpointTimer: ReturnType<typeof setTimeout> | null = null;
  private checkpointPromise: Promise<void> | null = null;

  /** exec_timeout_ms: per-chunk max execution time */
  constructor(
    rBin = 'Rscript',
    private execTimeoutMs = 0,
    sessionKey = 'default',
  ) {
    super();
    this.requestedBin = rBin;
    this.execTimeoutMs = sanitizeExecTimeoutMs(execTimeoutMs);
    this.kernelPath = path.join(__dirname, '..', 'r', 'kernel.R');
    const checkpointId = crypto.createHash('sha1').update(sessionKey).digest('hex');
    this.checkpointPath = path.join(os.tmpdir(), 'r-notebook-checkpoints', `${checkpointId}.rds`);
  }

  // ---- Lifecycle -----------------------------------------------------------

  async start(): Promise<void> {
    if (this.recoveryPromise) {
      await this.recoveryPromise;
      return;
    }
    if (!this.started && this.hasCheckpoint) {
      await this.beginRecoveryFromCheckpoint();
      return;
    }
    await this.startProcess();
  }

  /** Stop the active R run and keep the current session alive. */
  interrupt(): void {
    if (!this.proc || !this.activeExecKey) return;
    const now = Date.now();
    if (this.interruptingExecKey === this.activeExecKey && now - this.lastInterruptAt < 500) return;
    this.interruptingExecKey = this.activeExecKey;
    this.lastInterruptAt = now;
    this.interruptAttempts = 1;
    this.signalRunningProcess('SIGINT');
    this.scheduleInterruptRetry(this.activeExecKey);
  }

  async stop(): Promise<void> {
    this.clearScheduledCheckpoint();
    this.lastVars = { type: 'vars_result', vars: [] };
    this.workspaceDirty = false;
    this.hasCheckpoint = false;
    this.checkpointingDisabled = false;
    this.deleteCheckpointFile();
    await this.stopProcess(new Error('Session stopped'));
  }

  async restart(): Promise<void> {
    this.clearScheduledCheckpoint();
    this.lastVars = { type: 'vars_result', vars: [] };
    this.workspaceDirty = false;
    this.hasCheckpoint = false;
    this.checkpointingDisabled = false;
    this.deleteCheckpointFile();
    await this.stopProcess(new Error('Session stopped'));
    await this.startProcess();
  }

  setExecutablePath(rBin: string): void {
    this.requestedBin = rBin;
  }

  executablePath(): string {
    return this.requestedBin;
  }

  setExecTimeoutMs(execTimeoutMs: number): void {
    this.execTimeoutMs = sanitizeExecTimeoutMs(execTimeoutMs);
  }

  cachedVars(): VarsResult {
    return this.lastVars;
  }

  isBusy(): boolean {
    return this.activeExecKey !== null;
  }

  private varsTimeoutMs(): number {
    return Math.max(60_000, Math.min(this.execTimeoutMs, 300_000));
  }

  private restoreTimeoutMs(): number {
    return Math.max(60_000, Math.min(this.execTimeoutMs, 3_000_000));
  }

  // ---- High-level API ------------------------------------------------------

  async ping(): Promise<void> {
    await this.start();
    await this.sendWait({ type: 'ping' }, '__ping__', 5_000);
  }

  async reset(): Promise<ExecResult> {
    await this.start();
    this.lastVars = { type: 'vars_result', vars: [] };
    this.workspaceDirty = false;
    this.hasCheckpoint = false;
    this.checkpointingDisabled = false;
    this.deleteCheckpointFile();
    return this.sendWait({ type: 'reset' }, '__reset__') as Promise<ExecResult>;
  }

  async exec(
    chunkId: string,
    code: string,
    opts?: { fig_width?: number; fig_height?: number; dpi?: number },
  ): Promise<ExecResult> {
    await this.start();
    const req: KernelRequest = {
      type: 'exec',
      chunk_id: chunkId,
      code,
      ...(opts?.fig_width  != null ? { fig_width:  opts.fig_width }  : {}),
      ...(opts?.fig_height != null ? { fig_height: opts.fig_height } : {}),
      ...(opts?.dpi        != null ? { dpi:        opts.dpi }        : {}),
    };
    this.activeExecKey = chunkId;
    this.interruptingExecKey = null;
    try {
      const result = await this.sendWait(req, chunkId, this.execTimeoutMs) as ExecResult;
      if (!result.error && !this.checkpointingDisabled) {
        this.workspaceDirty = true;
        this.scheduleWorkspaceCheckpoint();
      }
      return result;
    } finally {
      if (this.activeExecKey === chunkId) this.activeExecKey = null;
      if (this.interruptingExecKey === chunkId) this.clearInterruptState();
    }
  }

  async dfPage(
    chunkId: string, name: string, page: number, pageSize = 50
  ): Promise<DfDataResult> {
    await this.start();
    const req: KernelRequest = {
      type: 'df_page', chunk_id: chunkId, name, page, page_size: pageSize,
    };
    return this.sendWait(req, `df:${chunkId}:${name}:${page}`) as Promise<DfDataResult>;
  }

  async vars(): Promise<VarsResult> {
    if (this.activeExecKey) return this.lastVars;
    await this.start();
    const result = await this.sendWait({ type: 'vars' }, '__vars__', this.varsTimeoutMs()) as VarsResult;
    this.lastVars = result;
    return result;
  }

  async complete(chunkId: string, code: string, cursorPos: number): Promise<string[]> {
    if (this.activeExecKey) return [];
    await this.start();
    const requestChunkId = `${chunkId}:cmp:${Date.now().toString(36)}:${Math.random().toString(36).slice(2, 8)}`;
    const req: KernelRequest = {
      type: 'complete', chunk_id: requestChunkId, code, cursor_pos: cursorPos,
    };
    const res = await this.sendWait(req, `cmp:${requestChunkId}`) as any;
    return res.completions ?? [];
  }

  async keepAlive(): Promise<void> {
    if (this.activeExecKey) return;
    await this.start();
    if (this.workspaceDirty && !this.checkpointingDisabled) {
      await this.ensureWorkspaceCheckpoint();
    }
    await this.sendWait({ type: 'ping' }, '__ping__', 5_000);
  }

  // ---- Internal machinery --------------------------------------------------

  private async startProcess(): Promise<void> {
    if (this.started && this.activeBin === this.requestedBin) return;
    if (this.started) await this.stopProcess(new Error('Session restarted'));
    this.clearScheduledCheckpoint();

    const myId = ++this.procId;
    const spawnBin = this.requestedBin;

    this.proc = cp.spawn(spawnBin, ['--vanilla', '--slave', this.kernelPath], {
      detached: process.platform !== 'win32',
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    this.activeBin = spawnBin;

    this.proc.on('error', (err) => {
      if (this.procId !== myId) return;
      this.started = false;
      this.activeBin = undefined;
      this.proc = null;
      this.rl = null;
      this.emit('error', err);
      this.rejectAllPending(new Error(`Failed to start R: ${err.message}`));
    });

    this.proc.on('exit', (code, signal) => {
      if (this.procId !== myId) return;
      const wasExpected = this.expectedExitProcId === myId;
      const interruptedExecKey = this.interruptingExecKey;
      const exitError = interruptedExecKey
        ? new Error('Interrupted by user')
        : new Error(`R process exited (code=${code})`);

      this.started = false;
      this.activeBin = undefined;
      this.proc = null;
      this.rl = null;
      this.activeExecKey = null;
      this.clearInterruptState();
      if (this.expectedExitProcId === myId) this.expectedExitProcId = null;

      if (!wasExpected) {
        this.emit('exit', { code, signal });
      }
      this.rejectAllPending(exitError);
      if (interruptedExecKey) {
        void this.beginRecoveryFromCheckpoint();
      }
    });

    this.proc.stderr!.on('data', (d: Buffer) => {
      const text = d.toString();
      if (this.expectedExitProcId === myId && text.trim() === 'Execution halted') return;
      this.emit('stderr', text);
    });

    this.rl = readline.createInterface({ input: this.proc.stdout! });
    this.rl.on('line', (line: string) => this.handleLine(line));

    await this.sendWait({ type: 'ping' }, '__ping__', 5_000);
    this.started = true;
  }

  private async stopProcess(reason: Error): Promise<void> {
    this.clearScheduledCheckpoint();
    if (!this.proc) {
      this.started = false;
      this.activeBin = undefined;
      this.activeExecKey = null;
      this.clearInterruptState();
      this.rejectAllPending(reason);
      return;
    }

    const procToStop = this.proc;
    const procIdToStop = this.procId;
    this.expectedExitProcId = this.procId;
    this.signalRunningProcess('SIGTERM');
    setTimeout(() => {
      if (this.procId !== procIdToStop) return;
      this.forceTerminateProcessGroup(procToStop);
    }, 1500);
    this.proc = null;
    this.rl = null;
    this.activeBin = undefined;
    this.started = false;
    this.activeExecKey = null;
    this.clearInterruptState();
    this.rejectAllPending(reason);
  }

  private clearInterruptState(): void {
    if (this.interruptRetryTimer) {
      clearInterval(this.interruptRetryTimer);
      this.interruptRetryTimer = null;
    }
    if (this.interruptEscalationTimer) {
      clearTimeout(this.interruptEscalationTimer);
      this.interruptEscalationTimer = null;
    }
    this.interruptingExecKey = null;
    this.interruptAttempts = 0;
    this.lastInterruptAt = 0;
  }

  private scheduleInterruptRetry(chunkId: string): void {
    if (this.interruptRetryTimer) {
      clearInterval(this.interruptRetryTimer);
    }
    this.interruptRetryTimer = setInterval(() => {
      if (!this.proc || this.activeExecKey !== chunkId) {
        this.clearInterruptState();
        return;
      }
      if (this.interruptAttempts >= 12) {
        clearInterval(this.interruptRetryTimer!);
        this.interruptRetryTimer = null;
        return;
      }
      this.interruptAttempts += 1;
      this.lastInterruptAt = Date.now();
      this.signalRunningProcess('SIGINT');
    }, 1000);
    this.interruptEscalationTimer = setTimeout(() => {
      this.interruptEscalationTimer = null;
      if (!this.proc || this.activeExecKey !== chunkId) return;
      void this.recoverAfterForcedInterrupt(chunkId);
    }, 4000);
  }

  private async recoverAfterForcedInterrupt(chunkId: string): Promise<void> {
    if (!this.proc || this.activeExecKey !== chunkId) return;
    await this.stopProcess(new Error('Interrupted by user'));
    await this.beginRecoveryFromCheckpoint();
  }

  private async ensureWorkspaceCheckpoint(): Promise<void> {
    if (!this.workspaceDirty || this.checkpointingDisabled) return;
    if (this.activeExecKey || this.recoveryPromise) return;
    if (this.checkpointPromise) return this.checkpointPromise;

    const snapshotTask = (async () => {
      let snapshot: any;
      try {
        snapshot = await this.sendWait(
          { type: 'snapshot', checkpoint_path: this.checkpointPath },
          '__snapshot__',
          this.restoreTimeoutMs(),
        ) as any;
      } catch {
        this.disableCheckpointing();
        return;
      }

      if (snapshot?.had_state === false) {
        this.workspaceDirty = false;
        this.hasCheckpoint = false;
        this.checkpointingDisabled = false;
        this.deleteCheckpointFile();
        return;
      }

      if (snapshot?.captured) {
        this.workspaceDirty = false;
        this.hasCheckpoint = true;
        this.checkpointingDisabled = false;
        return;
      }

      this.disableCheckpointing();
    })();

    let trackedTask: Promise<void>;
    trackedTask = snapshotTask.finally(() => {
      if (this.checkpointPromise === trackedTask) {
        this.checkpointPromise = null;
      }
      if (this.workspaceDirty && !this.checkpointingDisabled && !this.activeExecKey && !this.recoveryPromise) {
        this.scheduleWorkspaceCheckpoint();
      }
    });
    this.checkpointPromise = trackedTask;
    return trackedTask;
  }

  private disableCheckpointing(): void {
    this.clearScheduledCheckpoint();
    this.workspaceDirty = false;
    this.hasCheckpoint = false;
    this.checkpointingDisabled = true;
    this.deleteCheckpointFile();
  }

  private scheduleWorkspaceCheckpoint(delayMs = IDLE_CHECKPOINT_DELAY_MS): void {
    if (!this.workspaceDirty || this.checkpointingDisabled) return;
    if (this.checkpointTimer) clearTimeout(this.checkpointTimer);
    this.checkpointTimer = setTimeout(() => {
      this.checkpointTimer = null;
      if (this.activeExecKey || this.recoveryPromise || !this.started) {
        this.scheduleWorkspaceCheckpoint(delayMs);
        return;
      }
      void this.ensureWorkspaceCheckpoint();
    }, delayMs);
  }

  private clearScheduledCheckpoint(): void {
    if (this.checkpointTimer) {
      clearTimeout(this.checkpointTimer);
      this.checkpointTimer = null;
    }
  }

  private deleteCheckpointFile(): void {
    try {
      fs.unlinkSync(this.checkpointPath);
    } catch {
      // Ignore missing checkpoint files.
    }
  }

  private beginRecoveryFromCheckpoint(): Promise<void> {
    if (this.recoveryPromise) return this.recoveryPromise;

    const recovery = (async () => {
      await this.startProcess();
      if (!this.hasCheckpoint) return;
      await this.sendWait(
        { type: 'restore_workspace', checkpoint_path: this.checkpointPath },
        '__restore_workspace__',
        this.restoreTimeoutMs(),
      );
    })();

    let trackedRecovery: Promise<void>;
    trackedRecovery = recovery.finally(() => {
      if (this.recoveryPromise === trackedRecovery) {
        this.recoveryPromise = null;
      }
    });
    this.recoveryPromise = trackedRecovery;
    void recovery.catch((err) => {
      this.emit('error', err instanceof Error ? err : new Error(String(err)));
    });
    return trackedRecovery;
  }

  private signalRunningProcess(signal: NodeJS.Signals): void {
    const proc = this.proc;
    if (!proc) return;
    // Send only to the direct child process. If the child is a wrapper script it
    // is responsible for forwarding the signal to R. Sending to the process group
    // (-pid) AS WELL causes a double-signal: the wrapper receives SIGINT twice,
    // forwards it twice, and the second interrupt arrives while R's first
    // interrupt handler is still executing — killing R before it can respond.
    try {
      proc.kill(signal);
    } catch {
      // Ignore signaling races with process shutdown.
    }
  }

  private forceTerminateProcessGroup(proc: cp.ChildProcess | null): void {
    if (!proc) return;
    if (process.platform !== 'win32' && typeof proc.pid === 'number') {
      try {
        process.kill(-proc.pid, 'SIGKILL');
        return;
      } catch {
        // Fall back to direct child kill below.
      }
    }
    try {
      proc.kill('SIGKILL');
    } catch {
      // Ignore signaling races with process shutdown.
    }
  }

  private handleLine(line: string): void {
    if (!line.trim()) return;
    let msg: KernelResponse | null = null;
    try {
      const parsed = JSON.parse(line);
      if (parsed && typeof parsed === 'object' && typeof (parsed as any).type === 'string') {
        msg = parsed as KernelResponse;
      }
    } catch {
      // Non-JSON lines are user stdout being mirrored for live streaming.
    }

    if (!msg) {
      if (this.activeExecKey) {
        this.emit('stream', {
          type: 'stream',
          chunk_id: this.activeExecKey,
          stream: 'stdout',
          text: `${line}\n`,
        });
      } else {
        this.emit('parse_error', line);
      }
      return;
    }

    this.emit('message', msg);

    if (msg.type === 'progress') {
      this.emit('progress', msg);
      return;
    }
    if (msg.type === 'stream') {
      this.emit('stream', msg);
      return;
    }
    if (msg.type === 'stream_output') {
      this.emit('stream_output', msg);
      return;
    }

    if (msg.type === 'pong') {
      const p = this.pending.get('__ping__');
      if (p) { clearTimeout(p.timeout); p.resolve(msg); this.pending.delete('__ping__'); }
      return;
    }
    if (msg.type === 'snapshot_result') {
      const p = this.pending.get('__snapshot__');
      if (p) { clearTimeout(p.timeout); p.resolve(msg); this.pending.delete('__snapshot__'); }
      return;
    }
    if (msg.type === 'workspace_restored') {
      const p = this.pending.get('__restore_workspace__');
      if (p) { clearTimeout(p.timeout); p.resolve(msg); this.pending.delete('__restore_workspace__'); }
      return;
    }
    if (msg.type === 'vars_result') {
      this.lastVars = msg;
      const p = this.pending.get('__vars__');
      if (p) { clearTimeout(p.timeout); p.resolve(msg); this.pending.delete('__vars__'); }
      return;
    }

    const key = this.pendingKey(msg);
    const p   = this.pending.get(key);
    if (p) {
      clearTimeout(p.timeout);
      if (msg.type === 'error') {
        p.reject(new Error(msg.message));
      } else {
        p.resolve(msg);
      }
      this.pending.delete(key);
    }
  }

  private pendingKey(msg: KernelResponse): string {
    if (msg.type === 'result')          return (msg as ExecResult).chunk_id;
    if (msg.type === 'df_data')         return `df:${(msg as DfDataResult).chunk_id}:${(msg as DfDataResult).name}:${(msg as DfDataResult).page}`;
    if (msg.type === 'complete_result') return `cmp:${(msg as any).chunk_id}`;
    if (msg.type === 'error')           return (msg as any).chunk_id ?? '';
    return '';
  }

  private sendWait(
    req: KernelRequest,
    key: string,
    timeoutMs = this.execTimeoutMs,
  ): Promise<KernelResponse> {
    return new Promise((resolve, reject) => {
      const stdin = this.proc?.stdin;
      if (!stdin) {
        reject(new Error('R session is not running'));
        return;
      }
      const timeout = timeoutMs > 0 ? setTimeout(() => {
        this.pending.delete(key);
        if (req.type === 'exec') this.interrupt();
        reject(new Error(`Kernel timeout for key=${key}`));
      }, timeoutMs) : undefined;

      this.pending.set(key, { resolve, reject, timeout });
      const json = JSON.stringify(req) + '\n';
      stdin.write(json, (err) => {
        if (!err) return;
        const pending = this.pending.get(key);
        if (!pending) return;
        clearTimeout(pending.timeout);
        this.pending.delete(key);
        pending.reject(err instanceof Error ? err : new Error(String(err)));
      });
    });
  }

  private rejectAllPending(err: Error): void {
    for (const [, p] of this.pending) {
      clearTimeout(p.timeout);
      p.reject(err);
    }
    this.pending.clear();
  }
}

// ---------------------------------------------------------------------------
// Session registry — one session per document URI

const sessions = new Map<string, RSession>();

function checkpointPathForDocUri(docUri: string): string {
  const checkpointId = crypto.createHash('sha1').update(docUri).digest('hex');
  return path.join(os.tmpdir(), 'r-notebook-checkpoints', `${checkpointId}.rds`);
}

export function getOrCreateSession(
  docUri: string,
  rBin?: string,
  execTimeoutMs?: number,
): RSession {
  let s = sessions.get(docUri);
  if (!s) {
    s = new RSession(rBin, execTimeoutMs, docUri);
    sessions.set(docUri, s);
  } else {
    if (rBin) s.setExecutablePath(rBin);
    if (execTimeoutMs != null) s.setExecTimeoutMs(execTimeoutMs);
  }
  return s;
}

export async function disposeSession(docUri: string): Promise<void> {
  const s = sessions.get(docUri);
  if (s) {
    await s.stop();
    sessions.delete(docUri);
  }
}

export async function purgeSessionState(docUri: string): Promise<void> {
  await disposeSession(docUri);
  try {
    fs.unlinkSync(checkpointPathForDocUri(docUri));
  } catch {
    // Ignore missing checkpoint files.
  }
}

export function getSession(docUri: string): RSession | undefined {
  return sessions.get(docUri);
}

export function getAllSessions(): Map<string, RSession> {
  return sessions;
}
