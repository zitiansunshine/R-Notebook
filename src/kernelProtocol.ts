// =============================================================================
// kernelProtocol.ts — TypeScript types for the R Notebook kernel protocol
// =============================================================================

// ---- Outbound (host → R kernel) -------------------------------------------

export interface ExecMessage {
  type: 'exec';
  chunk_id: string;
  code: string;
  /** inches – maps to fig.width knitr option */
  fig_width?:  number;
  /** inches – maps to fig.height knitr option */
  fig_height?: number;
  /** dots per inch – maps to dpi knitr option */
  dpi?:        number;
}

export interface DfPageMessage {
  type: 'df_page';
  chunk_id: string;
  name: string;
  page: number;
  page_size?: number;
}

export interface CompleteMessage {
  type: 'complete';
  chunk_id: string;
  code: string;
  cursor_pos: number;
}

export interface PingMessage     { type: 'ping' }
export interface ResetMessage    { type: 'reset' }
export interface SnapshotMessage {
  type: 'snapshot';
  checkpoint_path?: string;
}
export interface RestoreWorkspaceMessage {
  type: 'restore_workspace';
  workspace_state?: string | null;
  checkpoint_path?: string | null;
}
export interface VarsMessage     { type: 'vars' }

export type KernelRequest =
  | ExecMessage
  | DfPageMessage
  | CompleteMessage
  | PingMessage
  | ResetMessage
  | SnapshotMessage
  | RestoreWorkspaceMessage
  | VarsMessage;

// ---- Inbound (R kernel → host) --------------------------------------------

export interface ColumnMeta {
  name: string;
  type: string;
}

export interface DataFrameResult {
  name: string;
  nrow: number;
  ncol: number;
  pages: number;
  page: number;
  row_names?: string[];
  columns: ColumnMeta[];
  /** rows × cols, values already cast to string/number/null */
  data: (string | number | null)[][];
}

export interface ConsoleSegment {
  code: string;
  output?: string;
}

export interface ExecResult {
  type: 'result';
  chunk_id: string;
  /** source code that produced this output, used for console echo */
  source_code?: string;
  /** per-expression console transcript in natural execution order */
  console_segments?: ConsoleSegment[];
  /** ordered combined console stream */
  console?: string;
  stdout: string;
  stderr: string;
  /** base64-encoded PNG strings (all plots, indexed by plot_idx from output_order) */
  plots: string[];
  /** raw HTML strings for interactive figures (e.g., Plotly) */
  plots_html?: string[];
  /** serialized data.frames, indexed by df index from output_order */
  dataframes: DataFrameResult[];
  /**
   * Natural interleaved order of plots and dataframes from the execution.
   * Each item references an index into plots[] or dataframes[].
   * When present, used instead of dataframes[]/plots[] for tab construction.
   */
  output_order?: Array<{ type: 'df' | 'plot'; index: number; name?: string }>;
  error: string | null;
}

export interface DfDataResult extends DataFrameResult {
  type: 'df_data';
  chunk_id: string;
}

export interface CompleteResult {
  type: 'complete_result';
  chunk_id: string;
  completions: string[];
}

export interface SnapshotResult {
  type: 'snapshot_result';
  names: string[];
  workspace_state?: string | null;
  captured?: boolean;
  had_state?: boolean;
}

export interface WorkspaceRestoredResult {
  type: 'workspace_restored';
}

export interface VarInfo {
  name:  string;
  type:  string;
  size:  string;
  value: string;
}

export interface VarsResult {
  type: 'vars_result';
  vars: VarInfo[];
}

export interface PongMessage { type: 'pong' }

export interface KernelError {
  type: 'error';
  chunk_id: string;
  message: string;
}

export interface ProgressMessage {
  type: 'progress';
  chunk_id: string;
  line: number;
  total: number;
  expr_code?: string;
}

export interface StreamMessage {
  type: 'stream';
  chunk_id: string;
  stream: 'stdout' | 'stderr';
  text: string;
}

export interface StreamOutputMessage {
  type: 'stream_output';
  chunk_id: string;
  kind: 'plot' | 'df';
  index: number;
  name?: string;
  b64?: string;           // present when kind === 'plot'
  df?: DataFrameResult;   // present when kind === 'df'
}

export type KernelResponse =
  | ExecResult
  | DfDataResult
  | CompleteResult
  | SnapshotResult
  | WorkspaceRestoredResult
  | VarsResult
  | PongMessage
  | KernelError
  | ProgressMessage
  | StreamMessage
  | StreamOutputMessage;
