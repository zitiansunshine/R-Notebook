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
