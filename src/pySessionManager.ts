// =============================================================================
// pySessionManager.ts — manages a persistent Python subprocess per document.
// Mirrors rSessionManager.ts; uses the same JSON-RPC protocol so the same
// kernelProtocol.ts types and notebookOutputHtml.ts builders are shared.
// =============================================================================

import * as cp       from 'child_process';
import * as path     from 'path';
import * as readline from 'readline';
import { EventEmitter } from 'events';
import { ExecResult } from './kernelProtocol';

interface Pending {
  resolve: (r: any) => void;
  reject:  (e: Error) => void;
  timeout: ReturnType<typeof setTimeout>;
}

function sanitizeExecTimeoutMs(timeoutMs?: number): number {
  return typeof timeoutMs === 'number' && Number.isFinite(timeoutMs) && timeoutMs > 0
    ? timeoutMs
    : 3_000_000;
}

export class PySession extends EventEmitter {
  private proc:    cp.ChildProcess | null = null;
  private rl:      readline.Interface | null = null;
  private pending: Map<string, Pending> = new Map();
  private started  = false;
  private procId   = 0;
  private readonly kernelPath: string;
  private requestedBin: string;
  private requestedEnv?: NodeJS.ProcessEnv;
  private activeSignature?: string;
  private activeExecKey: string | null = null;
  private lastVars: { type: 'vars_result'; vars: any[] } = { type: 'vars_result', vars: [] };

  constructor(
    pyBin = 'python3',
    private execTimeoutMs = 3_000_000,
    env?: NodeJS.ProcessEnv,
  ) {
    super();
    this.requestedBin = pyBin;
    this.requestedEnv = env;
    this.execTimeoutMs = sanitizeExecTimeoutMs(execTimeoutMs);
    this.kernelPath = path.join(__dirname, '..', 'python', 'kernel.py');
  }

  // ---- Lifecycle -----------------------------------------------------------

  async start(): Promise<void> {
    const requestedSignature = this.kernelSignature(this.requestedBin, this.requestedEnv);
    if (this.started && this.activeSignature === requestedSignature) return;
    if (this.started) await this.stop();

    const myId = ++this.procId;
    const spawnBin = this.requestedBin;
    const spawnEnv = this.requestedEnv ? { ...process.env, ...this.requestedEnv } : process.env;

    this.proc = cp.spawn(spawnBin, [this.kernelPath], {
      env: spawnEnv,
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    this.activeSignature = requestedSignature;

    this.proc.on('error', (err) => {
      if (this.procId !== myId) return;
      this.started = false;
      this.activeSignature = undefined;
      this.proc = null; this.rl = null;
      this.emit('error', err);
      this.rejectAll(new Error(`Failed to start Python: ${err.message}`));
    });

    this.proc.on('exit', (code, signal) => {
      if (this.procId !== myId) return;
      this.started = false;
      this.activeSignature = undefined;
      this.proc = null; this.rl = null;
      this.emit('exit', { code, signal });
      this.rejectAll(new Error(`Python process exited (code=${code})`));
    });

    this.proc.stderr!.on('data', (d: Buffer) => this.emit('stderr', d.toString()));

    this.rl = readline.createInterface({ input: this.proc.stdout! });
    this.rl.on('line', (line: string) => this.handleLine(line));

    await this.ping();
    this.started = true;
  }

  interrupt(): void { this.proc?.kill('SIGINT'); }

  async stop(): Promise<void> {
    this.lastVars = { type: 'vars_result', vars: [] };
    this.proc?.kill('SIGTERM');
    this.proc = null; this.rl = null; this.started = false; this.activeSignature = undefined;
    this.activeExecKey = null;
    this.rejectAll(new Error('Session stopped'));
  }

  async restart(): Promise<void> { await this.stop(); await this.start(); }

  setKernel(pyBin: string, env?: NodeJS.ProcessEnv): void {
    this.requestedBin = pyBin;
    this.requestedEnv = env;
  }

  setExecutablePath(pyBin: string): void {
    this.setKernel(pyBin, undefined);
  }

  executablePath(): string {
    return this.requestedBin;
  }

  setExecTimeoutMs(execTimeoutMs: number): void {
    this.execTimeoutMs = sanitizeExecTimeoutMs(execTimeoutMs);
  }

  cachedVars(): { type: 'vars_result'; vars: any[] } {
    return this.lastVars;
  }

  isBusy(): boolean {
    return this.activeExecKey !== null;
  }

  private varsTimeoutMs(): number {
    return Math.max(60_000, Math.min(this.execTimeoutMs, 300_000));
  }

  private kernelSignature(pyBin: string, env?: NodeJS.ProcessEnv): string {
    return JSON.stringify({
      pyBin,
      env: env ? Object.entries(env).sort(([a], [b]) => a.localeCompare(b)) : [],
    });
  }

  // ---- API -----------------------------------------------------------------

  async ping(): Promise<void> {
    await this.sendWait({ type: 'ping' }, '__ping__', 15_000);
  }

  async exec(
    chunkId: string,
    code: string,
    opts?: { fig_width?: number; fig_height?: number; dpi?: number },
  ): Promise<ExecResult> {
    const req: any = {
      type: 'exec', chunk_id: chunkId, code,
      ...(opts?.fig_width  != null ? { fig_width:  opts.fig_width  } : {}),
      ...(opts?.fig_height != null ? { fig_height: opts.fig_height } : {}),
      ...(opts?.dpi        != null ? { dpi:        opts.dpi        } : {}),
    };
    this.activeExecKey = chunkId;
    try {
      return await this.sendWait(req, chunkId, this.execTimeoutMs);
    } finally {
      if (this.activeExecKey === chunkId) this.activeExecKey = null;
    }
  }

  async vars(): Promise<{ type: 'vars_result'; vars: any[] }> {
    if (this.activeExecKey) return this.lastVars;
    const result = await this.sendWait({ type: 'vars' }, '__vars__', this.varsTimeoutMs());
    this.lastVars = result;
    return result;
  }

  // ---- Internal machinery --------------------------------------------------

  private handleLine(line: string): void {
    if (!line.trim()) return;
    let msg: any;
    try { msg = JSON.parse(line); } catch { return; }

    if (msg.type === 'pong') {
      const p = this.pending.get('__ping__');
      if (p) { clearTimeout(p.timeout); p.resolve(msg); this.pending.delete('__ping__'); }
      return;
    }
    if (msg.type === 'vars_result') {
      this.lastVars = msg;
      const p = this.pending.get('__vars__');
      if (p) { clearTimeout(p.timeout); p.resolve(msg); this.pending.delete('__vars__'); }
      return;
    }
    if (msg.type === 'stream') {
      this.emit('stream', msg);
      return;
    }

    const key = msg.chunk_id ?? '';
    const p   = this.pending.get(key);
    if (p) {
      clearTimeout(p.timeout);
      if (msg.type === 'error') {
        p.reject(new Error(msg.message ?? 'Python kernel error'));
      } else {
        p.resolve(msg);
      }
      this.pending.delete(key);
    }
  }

  private sendWait(req: any, key: string, timeoutMs = this.execTimeoutMs): Promise<any> {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(key);
        if (req.type === 'exec') this.interrupt();
        reject(new Error(`Python kernel timeout (key=${key})`));
      }, timeoutMs);
      this.pending.set(key, { resolve, reject, timeout });
      this.proc?.stdin?.write(JSON.stringify(req) + '\n');
    });
  }

  private rejectAll(err: Error): void {
    for (const [, p] of this.pending) { clearTimeout(p.timeout); p.reject(err); }
    this.pending.clear();
  }
}

// ---------------------------------------------------------------------------
// Session registry — one session per document URI

const sessions = new Map<string, PySession>();

export function getOrCreatePySession(
  docUri: string,
  pyBin?: string,
  env?: NodeJS.ProcessEnv,
  execTimeoutMs?: number,
): PySession {
  let s = sessions.get(docUri);
  if (!s) {
    s = new PySession(pyBin, execTimeoutMs, env);
    sessions.set(docUri, s);
  } else {
    if (pyBin) s.setKernel(pyBin, env);
    if (execTimeoutMs != null) s.setExecTimeoutMs(execTimeoutMs);
  }
  return s;
}

export function getPySession(docUri: string): PySession | undefined {
  return sessions.get(docUri);
}

export async function disposePySession(docUri: string): Promise<void> {
  const s = sessions.get(docUri);
  if (s) { await s.stop(); sessions.delete(docUri); }
}

export function getAllPySessions(): Map<string, PySession> { return sessions; }
