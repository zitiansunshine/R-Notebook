import * as vscode from 'vscode';

export const EXTENSION_BRAND = 'R Notebook';
export const NOTEBOOK_TYPE = 'rNotebook.rmarkdownNotebook';
export const PY_NOTEBOOK_TYPE = 'rNotebook.pythonNotebook';
export const RMARKDOWN_EDITOR_VIEW_TYPE = 'rNotebook.rmarkdownEditor';
export const VARIABLES_PANEL_VIEW_TYPE = 'rNotebook.variablesPanel';

export const COMMAND_IDS = {
  rRunChunk: 'rNotebook.r.runChunk',
  rRunAll: 'rNotebook.r.runAll',
  rRestartSession: 'rNotebook.r.restartSession',
  rSetCellFigureOptions: 'rNotebook.r.setCellFigureOptions',
  pySetCellFigureOptions: 'rNotebook.py.setCellFigureOptions',
  notebookSelectAvailableRKernels: 'rNotebook.notebook.selectAvailableRKernels',
  notebookSelectAvailablePythonKernels: 'rNotebook.notebook.selectAvailablePythonKernels',
  notebookInterrupt: 'rNotebook.notebook.interrupt',
  notebookRestart: 'rNotebook.notebook.restart',
  notebookShowVariables: 'rNotebook.notebook.showVariables',
  notebookClearOutputs: 'rNotebook.notebook.clearOutputs',
  notebookExport: 'rNotebook.notebook.export',
} as const;

const R_CONFIG_SECTION = 'rNotebook.r';
const LEGACY_R_CONFIG_SECTION = 'rnotebook.r';
const PYTHON_CONFIG_SECTION = 'rNotebook.python';
const LEGACY_PYTHON_CONFIG_SECTION = 'rnotebook.python';

function getConfigValue<T>(
  section: string,
  legacySection: string,
  key: string,
  defaultValue: T,
): T {
  const value = vscode.workspace.getConfiguration(section).get<T>(key);
  if (value !== undefined) return value;
  return vscode.workspace.getConfiguration(legacySection).get<T>(key, defaultValue)!;
}

export function getRConfigValue<T>(key: string, defaultValue: T): T {
  return getConfigValue(R_CONFIG_SECTION, LEGACY_R_CONFIG_SECTION, key, defaultValue);
}

export function getPythonConfigValue<T>(key: string, defaultValue: T): T {
  return getConfigValue(PYTHON_CONFIG_SECTION, LEGACY_PYTHON_CONFIG_SECTION, key, defaultValue);
}

export async function updateRConfigValue<T>(
  key: string,
  value: T,
  target: vscode.ConfigurationTarget,
): Promise<void> {
  await vscode.workspace.getConfiguration(R_CONFIG_SECTION).update(key, value, target);
}

export type RFigureOptions = {
  fig_width: number;
  fig_height: number;
  dpi: number;
};

export function getRAdditionalExecutablePaths(): string[] {
  return normalizePathList(getRConfigValue<string[]>('additionalRPaths', []));
}

export function getRExecutablePathCandidates(): string[] {
  return uniqueStrings([
    getRConfigValue('rPath', 'Rscript'),
    ...getRAdditionalExecutablePaths(),
  ]);
}

export async function rememberRExecutablePath(rPath: string): Promise<void> {
  const trimmed = rPath.trim();
  if (!trimmed) return;

  const defaultPath = getRConfigValue('rPath', 'Rscript').trim();
  if (trimmed === defaultPath) return;

  const current = getRAdditionalExecutablePaths();
  if (current.includes(trimmed)) return;
  await updateRConfigValue('additionalRPaths', [...current, trimmed], vscode.ConfigurationTarget.Global);
}

export function getRFigureDefaults(): RFigureOptions {
  return {
    fig_width: numberSetting('defaultFigWidth', 7, 1, 20),
    fig_height: numberSetting('defaultFigHeight', 5, 1, 20),
    dpi: numberSetting('defaultDpi', 120, 72, 600),
  };
}

export function mergeRFigureOptions(options?: Record<string, unknown>): RFigureOptions {
  const defaults = getRFigureDefaults();
  return {
    fig_width: numberOption(options?.fig_width, defaults.fig_width, 1, 20),
    fig_height: numberOption(options?.fig_height, defaults.fig_height, 1, 20),
    dpi: numberOption(options?.dpi, defaults.dpi, 72, 600),
  };
}

export async function updatePythonConfigValue<T>(
  key: string,
  value: T,
  target: vscode.ConfigurationTarget,
): Promise<void> {
  await vscode.workspace.getConfiguration(PYTHON_CONFIG_SECTION).update(key, value, target);
}

export function affectsPythonConfig(
  event: vscode.ConfigurationChangeEvent,
  key: string,
): boolean {
  return event.affectsConfiguration(`${PYTHON_CONFIG_SECTION}.${key}`)
    || event.affectsConfiguration(`${LEGACY_PYTHON_CONFIG_SECTION}.${key}`);
}

export function affectsRConfig(
  event: vscode.ConfigurationChangeEvent,
  key: string,
): boolean {
  return event.affectsConfiguration(`${R_CONFIG_SECTION}.${key}`)
    || event.affectsConfiguration(`${LEGACY_R_CONFIG_SECTION}.${key}`);
}

function normalizePathList(value: readonly string[] | undefined): string[] {
  return uniqueStrings((value ?? []).map((item) => String(item ?? '').trim()));
}

function uniqueStrings(values: readonly string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const value of values) {
    const trimmed = String(value ?? '').trim();
    if (!trimmed || seen.has(trimmed)) continue;
    seen.add(trimmed);
    result.push(trimmed);
  }
  return result;
}

function numberSetting(key: string, fallback: number, min: number, max: number): number {
  return numberOption(getRConfigValue<number>(key, fallback), fallback, min, max);
}

function numberOption(value: unknown, fallback: number, min: number, max: number): number {
  const numeric = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.min(max, Math.max(min, numeric));
}

export type NotebookKernelMetadata = {
  pythonPath?: string;
  rPath?: string;
  controllerId?: string;
  kernelspecName?: string;
  displayName?: string;
};

export function getNotebookKernelMetadata(meta: {
  rNotebook?: NotebookKernelMetadata;
  rnotebook?: NotebookKernelMetadata;
}): NotebookKernelMetadata | undefined {
  return meta.rNotebook ?? meta.rnotebook;
}
