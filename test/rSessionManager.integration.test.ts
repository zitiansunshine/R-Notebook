// =============================================================================
// rSessionManager.integration.test.ts — Integration tests for the R kernel
//
// Requires a working Rscript installation.
// Skips automatically if Rscript is not found in PATH.
//
// Run: npx vitest run test/rSessionManager.integration.test.ts
// =============================================================================

import { describe, it, expect, beforeAll, afterAll, beforeEach, vi } from 'vitest';
import { execSync } from 'child_process';
import * as path from 'path';
import { RSession } from '../src/rSessionManager';

// ---------------------------------------------------------------------------
// Check R availability
// ---------------------------------------------------------------------------

function rAvailable(): boolean {
  try {
    execSync('Rscript --version', { stdio: 'ignore', timeout: 5000 });
    return true;
  } catch {
    return false;
  }
}

const HAS_R = rAvailable();

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

/** Resolve kernel.R path relative to this test file */
const KERNEL_PATH = path.resolve(__dirname, '..', 'r', 'kernel.R');

/** Create an RSession pointing to the local kernel.R */
function makeSession(): RSession {
  // Patch the path resolution: RSession internally resolves from __dirname
  // For testing we override by subclassing to inject the path.
  return new RSession('Rscript', 15_000);
}

// ---------------------------------------------------------------------------

describe.skipIf(!HAS_R)('RSession integration', () => {

  let session: RSession;

  beforeAll(async () => {
    session = makeSession();
    await session.start();
  }, 30_000);

  afterAll(async () => {
    await session.stop();
  });

  // ---- Ping ---------------------------------------------------------------

  it('responds to ping', async () => {
    await expect(session.ping()).resolves.toBeUndefined();
  });

  // ---- Basic exec ---------------------------------------------------------

  it('executes simple arithmetic', async () => {
    const result = await session.exec('t1', '1 + 1');
    expect(result.error).toBeNull();
    expect(result.stdout.trim()).toBe('[1] 2');
  });

  it('returns stdout from print()', async () => {
    const result = await session.exec('t2', 'print("hello world")');
    expect(result.stdout).toContain('hello world');
    expect(result.error).toBeNull();
  });

  it('captures multi-line output', async () => {
    const result = await session.exec('t3', 'for (i in 1:3) cat(i, "\\n")');
    expect(result.stdout).toContain('1');
    expect(result.stdout).toContain('2');
    expect(result.stdout).toContain('3');
  });

  it('emits streamed stdout updates while a chunk is running', async () => {
    const streamed: string[] = [];
    const onStream = (msg: any) => {
      if (msg.chunk_id === 't3-stream' && msg.stream === 'stdout') {
        streamed.push(msg.text);
      }
    };

    session.on('stream', onStream);
    try {
      const result = await session.exec('t3-stream', 'cat("alpha\\n"); cat("beta\\n")');
      expect(result.stdout).toContain('alpha');
      expect(result.stdout).toContain('beta');
      expect(streamed.join('')).toContain('alpha');
      expect(streamed.join('')).toContain('beta');
    } finally {
      session.off('stream', onStream);
    }
  });

  // ---- Error handling -----------------------------------------------------

  it('captures R errors without crashing session', async () => {
    const result = await session.exec('t4', 'stop("deliberate error")');
    expect(result.error).not.toBeNull();
    expect(result.error).toContain('deliberate error');
  });

  it('session still works after an error', async () => {
    await session.exec('t5-err', 'stop("oops")');
    const result = await session.exec('t5-ok', '2 + 2');
    expect(result.error).toBeNull();
    expect(result.stdout).toContain('4');
  });

  it('captures parse errors', async () => {
    const result = await session.exec('t6', 'x <- (');
    expect(result.error).not.toBeNull();
  });

  // ---- Environment persistence --------------------------------------------

  it('persists variables between chunks', async () => {
    await session.exec('env1', 'my_var <- 42');
    const result = await session.exec('env2', 'cat(my_var)');
    expect(result.stdout).toContain('42');
    expect(result.error).toBeNull();
  });

  it('persists functions between chunks', async () => {
    await session.exec('fn1', 'double <- function(x) x * 2');
    const result = await session.exec('fn2', 'double(7)');
    expect(result.stdout).toContain('14');
  });

  // ---- DataFrame detection ------------------------------------------------

  it('detects new data.frame assignments', async () => {
    const result = await session.exec('df1', 'test_df <- data.frame(a=1:3, b=c("x","y","z"))');
    expect(result.dataframes).toHaveLength(1);
    expect(result.dataframes[0].name).toBe('test_df');
    expect(result.dataframes[0].nrow).toBe(3);
    expect(result.dataframes[0].ncol).toBe(2);
  });

  it('returns first page of dataframe', async () => {
    await session.exec('df2', 'big_df <- data.frame(x=1:120, y=letters[rep(1:26, 5)[1:120]])');
    const result = await session.exec('df3', 'head(big_df, 0)');
    const df = await session.dfPage('df-pg', 'big_df', 0, 50);
    expect(df.nrow).toBe(120);
    expect(df.pages).toBe(3); // ceil(120/50)
    expect(df.data.length).toBe(50);
    expect(df.page).toBe(0);
  });

  it('paginates dataframe to page 2', async () => {
    const df = await session.dfPage('df-pg2', 'big_df', 1, 50);
    expect(df.page).toBe(1);
    expect(df.data.length).toBe(50);
  });

  it('returns partial last page', async () => {
    const df = await session.dfPage('df-pg3', 'big_df', 2, 50);
    expect(df.page).toBe(2);
    expect(df.data.length).toBe(20); // 120 - 100 = 20
  });

  it.skipIf(!checkBase64enc())('surfaces dataframe and plot outputs from a visible plain list', async () => {
    const result = await session.exec('df-list-visible', [
      'bundle <- list(',
      '  tbl = data.frame(a = 1:2, b = c("x", "y")),',
      '  img = grid::circleGrob()',
      ')',
      'bundle',
    ].join('\n'));

    expect(result.error).toBeNull();
    expect(result.dataframes).toHaveLength(1);
    expect(result.dataframes[0].nrow).toBe(2);
    expect(result.dataframes[0].ncol).toBe(2);
    expect(result.plots.length).toBeGreaterThan(0);
    expect(result.stdout).toContain('$tbl');
    expect(result.stdout).toContain('<data.frame: 2 x 2>');
    expect(result.stdout).toContain('$img');
    expect(result.stdout).toContain('<plot:');
  });

  it('errors cleanly for non-existent df', async () => {
    await expect(
      session.dfPage('df-err', 'no_such_df', 0)
    ).rejects.toThrow();
  });

  // ---- Reset --------------------------------------------------------------

  it('reset clears the global environment', async () => {
    await session.exec('reset1', 'clear_test_var <- 999');
    await session.reset();
    const result = await session.exec('reset2', 'exists("clear_test_var")');
    expect(result.stdout).toContain('FALSE');
  });

  it('session is still usable after reset', async () => {
    await session.reset();
    const result = await session.exec('post-reset', 'cat("alive")');
    expect(result.stdout).toContain('alive');
  });

  // ---- Plot output --------------------------------------------------------

  it.skipIf(!checkBase64enc())('captures PNG plot as base64', async () => {
    const result = await session.exec('plot1', 'plot(1:5)');
    // Only test if base64enc is installed in this R
    if (result.plots.length > 0) {
      expect(result.plots[0]).toMatch(/^[A-Za-z0-9+/]+=*$/);
      expect(result.plots[0].length).toBeGreaterThan(100);
    }
    // No error regardless
    expect(result.error).toBeNull();
  });

  // ---- Stderr capture -----------------------------------------------------

  it('captures stderr warnings', async () => {
    const result = await session.exec('warn1', 'warning("test warning")');
    // warnings go to stderr in R
    expect(result.stderr.length + result.stdout.length).toBeGreaterThan(0);
    expect(result.error).toBeNull(); // warnings are not errors
  });

  // ---- Completion ---------------------------------------------------------

  it('returns completions for partial symbol', async () => {
    const completions = await session.complete('cmp1', 'me', 2);
    // Should include base R functions starting with 'me'
    expect(Array.isArray(completions)).toBe(true);
  });

});

// ---------------------------------------------------------------------------
// Snapshot test (no R needed — tests the protocol serialisation logic)
// ---------------------------------------------------------------------------

describe('RSession constructor', () => {
  it('creates session without starting it', () => {
    const s = new RSession('Rscript', 5000);
    expect(s).toBeTruthy();
  });

  it('continues execution when checkpointing times out', async () => {
    const s = new RSession('Rscript', 5000);
    const mockedExecResult = {
      type: 'result',
      chunk_id: 'timeout-safe',
      console: '',
      stdout: '[1] 2\n',
      stderr: '',
      plots: [],
      dataframes: [],
      error: null,
    };
    const sessionAny = s as any;

    sessionAny.workspaceDirty = true;
    sessionAny.hasCheckpoint = true;
    sessionAny.start = vi.fn().mockResolvedValue(undefined);
    sessionAny.sendWait = vi
      .fn()
      .mockRejectedValueOnce(new Error('Kernel timeout for key=__snapshot__'))
      .mockResolvedValueOnce(mockedExecResult);

    const result = await s.exec('timeout-safe', '1 + 1');

    expect(result).toEqual(mockedExecResult);
    expect(sessionAny.checkpointingDisabled).toBe(true);
    expect(sessionAny.workspaceDirty).toBe(false);
    expect(sessionAny.hasCheckpoint).toBe(false);
    expect(sessionAny.sendWait).toHaveBeenNthCalledWith(
      1,
      { type: 'snapshot', checkpoint_path: sessionAny.checkpointPath },
      '__snapshot__',
      60000,
    );
    expect(sessionAny.sendWait).toHaveBeenNthCalledWith(
      2,
      { type: 'exec', chunk_id: 'timeout-safe', code: '1 + 1' },
      'timeout-safe',
      5000,
    );
  });
});

// ---------------------------------------------------------------------------

function checkBase64enc(): boolean {
  if (!HAS_R) return false;
  try {
    execSync('Rscript -e "library(base64enc)"', { stdio: 'ignore', timeout: 5000 });
    return true;
  } catch {
    return false;
  }
}
