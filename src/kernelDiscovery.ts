import * as cp from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as vscode from 'vscode';

export type JupyterKernelSpec = {
  argv?: string[];
  display_name?: string;
  env?: Record<string, string>;
  language?: string;
};

export type JupyterKernelSpecEntry = {
  name: string;
  resourceDir?: string;
  spec: JupyterKernelSpec;
};

export type PythonKernelDescriptor = {
  id: string;
  label: string;
  description?: string;
  detail?: string;
  env?: NodeJS.ProcessEnv;
  environmentType?: 'conda' | 'venv' | 'pyenv' | 'system';
  pythonPath: string;
  displayName: string;
  kernelspecName: string;
  source: 'kernelspec' | 'executable';
};

export type RKernelDescriptor = {
  id: string;
  label: string;
  description?: string;
  detail?: string;
  rPath: string;
  displayName: string;
  kernelspecName?: string;
  source: 'kernelspec' | 'executable';
};

type RPickItem = vscode.QuickPickItem & {
  action?: 'browse' | 'manual';
  rPath?: string;
};

export function discoverPythonKernels(configuredPython: string): PythonKernelDescriptor[] {
  const descriptors: PythonKernelDescriptor[] = [];
  const seen = new Set<string>();
  const jupyterSpecs = queryJupyterKernelspecs();

  for (const entry of jupyterSpecs) {
    if (entry.spec.language?.toLowerCase() !== 'python') continue;

    const pythonPath = resolvePythonPath(entry.spec.argv?.[0]);
    const displayName = entry.spec.display_name ?? entry.name;
    const descriptor: PythonKernelDescriptor = {
      id: `kernelspec:${entry.name}:${pythonPath}`,
      label: displayName,
      description: 'Jupyter kernel',
      detail: pythonPath,
      env: entry.spec.env,
      pythonPath,
      displayName,
      kernelspecName: entry.name,
      source: 'kernelspec',
    };

    if (seen.has(descriptor.id)) continue;
    seen.add(descriptor.id);
    descriptors.push(descriptor);
  }

  for (const pythonPath of detectPythonInstallations(configuredPython, jupyterSpecs)) {
    const descriptor: PythonKernelDescriptor = {
      id: `executable:${pythonPath}`,
      label: fallbackPythonLabel(pythonPath),
      description: fallbackPythonDescription(pythonPath),
      detail: pythonPath,
      env: undefined,
      environmentType: detectPythonEnvironment(pythonPath).type,
      pythonPath,
      displayName: fallbackPythonLabel(pythonPath),
      kernelspecName: fallbackKernelName(pythonPath),
      source: 'executable',
    };

    if (seen.has(descriptor.id)) continue;
    seen.add(descriptor.id);
    descriptors.push(descriptor);
  }

  if (descriptors.length === 0) {
    const fallbackPath = configuredPython || 'python3';
    descriptors.push({
      id: `executable:${fallbackPath}`,
      label: fallbackPythonLabel(fallbackPath),
      description: fallbackPythonDescription(fallbackPath),
      detail: fallbackPath,
      env: undefined,
      environmentType: detectPythonEnvironment(fallbackPath).type,
      pythonPath: fallbackPath,
      displayName: fallbackPythonLabel(fallbackPath),
      kernelspecName: fallbackKernelName(fallbackPath),
      source: 'executable',
    });
  }

  return descriptors;
}

export async function discoverPythonKernelsAsync(configuredPython: string): Promise<PythonKernelDescriptor[]> {
  const descriptors: PythonKernelDescriptor[] = [];
  const seen = new Set<string>();
  const jupyterSpecs = await queryJupyterKernelspecsAsync();

  for (const entry of jupyterSpecs) {
    if (entry.spec.language?.toLowerCase() !== 'python') continue;

    const pythonPath = resolvePythonPath(entry.spec.argv?.[0]);
    const displayName = entry.spec.display_name ?? entry.name;
    const descriptor: PythonKernelDescriptor = {
      id: `kernelspec:${entry.name}:${pythonPath}`,
      label: displayName,
      description: 'Jupyter kernel',
      detail: pythonPath,
      env: entry.spec.env,
      pythonPath,
      displayName,
      kernelspecName: entry.name,
      source: 'kernelspec',
    };

    if (seen.has(descriptor.id)) continue;
    seen.add(descriptor.id);
    descriptors.push(descriptor);
  }

  for (const pythonPath of await detectPythonInstallationsAsync(configuredPython, jupyterSpecs)) {
    const info = await getPythonInterpreterInfoAsync(pythonPath);
    const label = formatPythonLabel(pythonPath, info);
    const descriptor: PythonKernelDescriptor = {
      id: `executable:${pythonPath}`,
      label,
      description: formatPythonDescription(info),
      detail: pythonPath,
      env: undefined,
      environmentType: info.type,
      pythonPath,
      displayName: label,
      kernelspecName: fallbackKernelName(pythonPath),
      source: 'executable',
    };

    if (seen.has(descriptor.id)) continue;
    seen.add(descriptor.id);
    descriptors.push(descriptor);
  }

  if (descriptors.length === 0) {
    const fallbackPath = configuredPython || 'python3';
    const info = await getPythonInterpreterInfoAsync(fallbackPath);
    const label = formatPythonLabel(fallbackPath, info);
    descriptors.push({
      id: `executable:${fallbackPath}`,
      label,
      description: formatPythonDescription(info),
      detail: fallbackPath,
      env: undefined,
      environmentType: info.type,
      pythonPath: fallbackPath,
      displayName: label,
      kernelspecName: fallbackKernelName(fallbackPath),
      source: 'executable',
    });
  }

  return descriptors;
}

export function discoverRKernels(
  configuredR: string,
  additionalRPaths: readonly string[] = [],
): RKernelDescriptor[] {
  const descriptors: RKernelDescriptor[] = [];
  const seen = new Set<string>();
  const jupyterSpecs = queryJupyterKernelspecs();

  for (const entry of jupyterSpecs) {
    if (entry.spec.language?.toLowerCase() !== 'r') continue;

    const rPath = resolveRScriptPath(entry.spec.argv?.[0]);
    if (!rPath) continue;

    const descriptor: RKernelDescriptor = {
      id: `kernelspec:${entry.name}:${rPath}`,
      label: entry.spec.display_name ?? entry.name,
      description: 'Jupyter kernel',
      detail: rPath,
      rPath,
      displayName: entry.spec.display_name ?? entry.name,
      kernelspecName: entry.name,
      source: 'kernelspec',
    };

    const key = `${descriptor.kernelspecName}:${descriptor.rPath}`;
    if (seen.has(key)) continue;
    seen.add(key);
    descriptors.push(descriptor);
  }

  for (const rPath of detectRInstallations(configuredR, additionalRPaths, jupyterSpecs)) {
    const key = `exec:${rPath}`;
    if (seen.has(key)) continue;
    seen.add(key);
    descriptors.push({
      id: `executable:${rPath}`,
      label: fallbackRLabel(rPath),
      description: 'R executable',
      detail: rPath,
      rPath,
      displayName: fallbackRLabel(rPath),
      source: 'executable',
    });
  }

  return descriptors;
}

export async function discoverRKernelsAsync(
  configuredR: string,
  additionalRPaths: readonly string[] = [],
): Promise<RKernelDescriptor[]> {
  const descriptors: RKernelDescriptor[] = [];
  const seen = new Set<string>();
  const jupyterSpecs = await queryJupyterKernelspecsAsync();

  for (const entry of jupyterSpecs) {
    if (entry.spec.language?.toLowerCase() !== 'r') continue;

    const rPath = resolveRScriptPath(entry.spec.argv?.[0]);
    if (!rPath) continue;

    const descriptor: RKernelDescriptor = {
      id: `kernelspec:${entry.name}:${rPath}`,
      label: entry.spec.display_name ?? entry.name,
      description: 'Jupyter kernel',
      detail: rPath,
      rPath,
      displayName: entry.spec.display_name ?? entry.name,
      kernelspecName: entry.name,
      source: 'kernelspec',
    };

    const key = `${descriptor.kernelspecName}:${descriptor.rPath}`;
    if (seen.has(key)) continue;
    seen.add(key);
    descriptors.push(descriptor);
  }

  for (const rPath of await detectRInstallationsAsync(configuredR, additionalRPaths, jupyterSpecs)) {
    const key = `exec:${rPath}`;
    if (seen.has(key)) continue;
    seen.add(key);
    descriptors.push({
      id: `executable:${rPath}`,
      label: fallbackRLabel(rPath),
      description: 'R executable',
      detail: rPath,
      rPath,
      displayName: fallbackRLabel(rPath),
      source: 'executable',
    });
  }

  return descriptors;
}

export async function pickRKernelPath(options?: {
  currentPath?: string;
  additionalPaths?: readonly string[];
  title?: string;
  placeHolder?: string;
}): Promise<string | undefined> {
  const currentPath = options?.currentPath ?? 'Rscript';
  const items: RPickItem[] = [
    ...(await discoverRKernelsAsync(currentPath, options?.additionalPaths ?? [])).map((descriptor) => ({
      label: descriptor.label,
      description: descriptor.rPath === currentPath
        ? `${descriptor.description ?? 'R executable'} · current`
        : descriptor.description,
      detail: descriptor.rPath,
      rPath: descriptor.rPath,
    })),
    {
      label: '$(folder-opened) Browse…',
      description: 'Pick Rscript from the file system',
      action: 'browse',
    },
    {
      label: '$(pencil) Enter manually…',
      description: 'Type the full path to Rscript',
      action: 'manual',
    },
  ];

  const pick = await vscode.window.showQuickPick(items, {
    title: options?.title ?? 'Select R Kernel (Rscript path)',
    placeHolder: options?.placeHolder ?? currentPath,
    matchOnDescription: true,
    matchOnDetail: true,
  });
  if (!pick) return undefined;

  if (pick.action === 'browse') {
    const uris = await vscode.window.showOpenDialog({
      title: 'Select Rscript',
      filters: { Executable: ['*'] },
    });
    return uris?.[0]?.fsPath;
  }

  if (pick.action === 'manual') {
    return vscode.window.showInputBox({
      title: 'R Kernel Path',
      prompt: 'Full path to Rscript (for example /usr/local/bin/Rscript)',
      value: currentPath,
    });
  }

  return pick.rPath;
}

export function fallbackKernelName(pythonPath: string): string {
  const base = path.basename(pythonPath).replace(/[^A-Za-z0-9_.-]+/g, '-');
  return base || 'python3';
}

export function fallbackPythonLabel(pythonPath: string): string {
  const info = getPythonInterpreterInfo(pythonPath);
  if (info.version && info.envName) return `Python ${info.version} (${info.envName})`;
  if (info.version) return `Python ${info.version}`;
  if (info.envName) return `Python (${info.envName})`;

  const base = path.basename(pythonPath) || pythonPath;
  return base === pythonPath ? `Python (${pythonPath})` : `Python (${base})`;
}

export function fallbackPythonDescription(pythonPath: string): string {
  const info = getPythonInterpreterInfo(pythonPath);
  switch (info.type) {
    case 'conda':
      return 'Conda environment';
    case 'venv':
      return 'Virtual environment';
    case 'pyenv':
      return 'pyenv environment';
    default:
      return 'Python environment';
  }
}

function fallbackRLabel(rPath: string): string {
  const base = path.basename(rPath) || rPath;
  return base === rPath ? `R (${rPath})` : `R (${base})`;
}

function detectPythonInstallations(
  configuredPython: string,
  jupyterSpecs: JupyterKernelSpecEntry[],
): string[] {
  const seen = new Set<string>();
  const candidates: string[] = [];

  function add(candidate?: string): void {
    if (!candidate || seen.has(candidate)) return;
    seen.add(candidate);
    candidates.push(candidate);
  }

  add(configuredPython);
  add('python3');
  add('python');

  for (const command of ['python3', 'python']) {
    for (const resolved of collectCommandCandidates(command)) add(resolved);
  }

  for (const candidate of [
    '/usr/local/bin/python3',
    '/opt/homebrew/bin/python3',
    '/usr/bin/python3',
    '/opt/homebrew/bin/python',
  ]) {
    if (pathExists(candidate)) add(candidate);
  }

  for (const candidate of collectWorkspacePythonCandidates()) add(candidate);
  for (const candidate of collectHomePythonCandidates()) add(candidate);

  for (const entry of jupyterSpecs) {
    if (entry.spec.language?.toLowerCase() !== 'python') continue;
    add(resolvePythonPath(entry.spec.argv?.[0]));
  }

  return candidates;
}

function detectRInstallations(
  configuredR: string,
  additionalRPaths: readonly string[],
  jupyterSpecs: JupyterKernelSpecEntry[],
): string[] {
  const seen = new Set<string>();
  const candidates: string[] = [];

  function add(candidate?: string): void {
    if (!candidate || seen.has(candidate)) return;
    seen.add(candidate);
    candidates.push(candidate);
  }

  add(configuredR);
  for (const candidate of additionalRPaths) add(candidate);
  add('Rscript');

  for (const resolved of collectCommandCandidates('Rscript')) add(resolved);

  for (const candidate of [
    '/usr/local/bin/Rscript',
    '/opt/homebrew/bin/Rscript',
    '/usr/bin/Rscript',
    '/opt/homebrew/opt/r/bin/Rscript',
  ]) {
    if (pathExists(candidate)) add(candidate);
  }

  for (const entry of jupyterSpecs) {
    if (entry.spec.language?.toLowerCase() !== 'r') continue;
    add(resolveRScriptPath(entry.spec.argv?.[0]));
  }

  return candidates;
}

async function detectPythonInstallationsAsync(
  configuredPython: string,
  jupyterSpecs: JupyterKernelSpecEntry[],
): Promise<string[]> {
  const seen = new Set<string>();
  const candidates: string[] = [];

  function add(candidate?: string): void {
    if (!candidate || seen.has(candidate)) return;
    seen.add(candidate);
    candidates.push(candidate);
  }

  add(configuredPython);
  add('python3');
  add('python');

  for (const command of ['python3', 'python']) {
    for (const resolved of await collectCommandCandidatesAsync(command)) add(resolved);
  }

  for (const candidate of [
    '/usr/local/bin/python3',
    '/opt/homebrew/bin/python3',
    '/usr/bin/python3',
    '/opt/homebrew/bin/python',
  ]) {
    if (await pathExistsAsync(candidate)) add(candidate);
  }

  for (const candidate of await collectWorkspacePythonCandidatesAsync()) add(candidate);
  for (const candidate of await collectHomePythonCandidatesAsync()) add(candidate);

  for (const entry of jupyterSpecs) {
    if (entry.spec.language?.toLowerCase() !== 'python') continue;
    add(resolvePythonPath(entry.spec.argv?.[0]));
  }

  return candidates;
}

async function detectRInstallationsAsync(
  configuredR: string,
  additionalRPaths: readonly string[],
  jupyterSpecs: JupyterKernelSpecEntry[],
): Promise<string[]> {
  const seen = new Set<string>();
  const candidates: string[] = [];

  function add(candidate?: string): void {
    if (!candidate || seen.has(candidate)) return;
    seen.add(candidate);
    candidates.push(candidate);
  }

  add(configuredR);
  for (const candidate of additionalRPaths) add(candidate);
  add('Rscript');

  for (const resolved of await collectCommandCandidatesAsync('Rscript')) add(resolved);

  for (const candidate of [
    '/usr/local/bin/Rscript',
    '/opt/homebrew/bin/Rscript',
    '/usr/bin/Rscript',
    '/opt/homebrew/opt/r/bin/Rscript',
  ]) {
    if (await pathExistsAsync(candidate)) add(candidate);
  }

  for (const entry of jupyterSpecs) {
    if (entry.spec.language?.toLowerCase() !== 'r') continue;
    add(resolveRScriptPath(entry.spec.argv?.[0]));
  }

  return candidates;
}

function collectWorkspacePythonCandidates(): string[] {
  const candidates: string[] = [];

  for (const folder of vscode.workspace.workspaceFolders ?? []) {
    const root = folder.uri.fsPath;
    for (const envName of ['.venv', 'venv', 'env']) {
      candidates.push(path.join(root, envName, 'bin', 'python'));
    }
  }

  return candidates.filter(pathExists);
}

async function collectWorkspacePythonCandidatesAsync(): Promise<string[]> {
  const candidates: string[] = [];

  for (const folder of vscode.workspace.workspaceFolders ?? []) {
    const root = folder.uri.fsPath;
    for (const envName of ['.venv', 'venv', 'env']) {
      candidates.push(path.join(root, envName, 'bin', 'python'));
    }
  }

  const existing = await Promise.all(candidates.map(async (candidate) =>
    (await pathExistsAsync(candidate)) ? candidate : undefined,
  ));
  return existing.filter((candidate): candidate is string => Boolean(candidate));
}

function collectHomePythonCandidates(): string[] {
  const home = os.homedir();
  const candidates: string[] = [
    process.env.CONDA_PREFIX ? path.join(process.env.CONDA_PREFIX, 'bin', 'python') : '',
    path.join(home, 'miniconda3', 'bin', 'python'),
    path.join(home, 'anaconda3', 'bin', 'python'),
    path.join(home, 'miniforge3', 'bin', 'python'),
    path.join(home, 'mambaforge', 'bin', 'python'),
    path.join(home, 'micromamba', 'bin', 'python'),
  ].filter(Boolean);

  for (const root of [
    path.join(home, '.virtualenvs'),
    path.join(home, '.pyenv', 'versions'),
    path.join(home, 'miniconda3', 'envs'),
    path.join(home, 'anaconda3', 'envs'),
    path.join(home, 'miniforge3', 'envs'),
    path.join(home, 'mambaforge', 'envs'),
    path.join(home, '.conda', 'envs'),
    path.join(home, 'micromamba', 'envs'),
  ]) {
    candidates.push(...collectExecutablesFromRoot(root, path.join('bin', 'python')));
  }

  return candidates.filter(pathExists);
}

async function collectHomePythonCandidatesAsync(): Promise<string[]> {
  const home = os.homedir();
  const candidates: string[] = [
    process.env.CONDA_PREFIX ? path.join(process.env.CONDA_PREFIX, 'bin', 'python') : '',
    path.join(home, 'miniconda3', 'bin', 'python'),
    path.join(home, 'anaconda3', 'bin', 'python'),
    path.join(home, 'miniforge3', 'bin', 'python'),
    path.join(home, 'mambaforge', 'bin', 'python'),
    path.join(home, 'micromamba', 'bin', 'python'),
  ].filter(Boolean);

  for (const root of [
    path.join(home, '.virtualenvs'),
    path.join(home, '.pyenv', 'versions'),
    path.join(home, 'miniconda3', 'envs'),
    path.join(home, 'anaconda3', 'envs'),
    path.join(home, 'miniforge3', 'envs'),
    path.join(home, 'mambaforge', 'envs'),
    path.join(home, '.conda', 'envs'),
    path.join(home, 'micromamba', 'envs'),
  ]) {
    candidates.push(...await collectExecutablesFromRootAsync(root, path.join('bin', 'python')));
  }

  const existing = await Promise.all(candidates.map(async (candidate) =>
    (await pathExistsAsync(candidate)) ? candidate : undefined,
  ));
  return existing.filter((candidate): candidate is string => Boolean(candidate));
}

function collectExecutablesFromRoot(root: string, suffix: string): string[] {
  if (!pathExists(root)) return [];

  try {
    return fs.readdirSync(root, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => path.join(root, entry.name, suffix))
      .filter(pathExists);
  } catch {
    return [];
  }
}

async function collectExecutablesFromRootAsync(root: string, suffix: string): Promise<string[]> {
  if (!(await pathExistsAsync(root))) return [];

  try {
    const entries = await fs.promises.readdir(root, { withFileTypes: true });
    const candidates = entries
      .filter((entry) => entry.isDirectory())
      .map((entry) => path.join(root, entry.name, suffix));
    const existing = await Promise.all(candidates.map(async (candidate) =>
      (await pathExistsAsync(candidate)) ? candidate : undefined,
    ));
    return existing.filter((candidate): candidate is string => Boolean(candidate));
  } catch {
    return [];
  }
}

function collectCommandCandidates(command: string): string[] {
  try {
    const result = cp.spawnSync(
      process.platform === 'win32' ? 'where' : 'which',
      process.platform === 'win32' ? [command] : ['-a', command],
      {
        encoding: 'utf-8',
        timeout: 3000,
        windowsHide: true,
      },
    );
    if (result.status !== 0 || !result.stdout) return [];
    return result.stdout
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
  } catch {
    return [];
  }
}

async function collectCommandCandidatesAsync(command: string): Promise<string[]> {
  const result = await execFileQuiet(
    process.platform === 'win32' ? 'where' : 'which',
    process.platform === 'win32' ? [command] : ['-a', command],
    3000,
  );
  if (!result || !result.stdout) return [];
  return result.stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

const pythonInterpreterInfoCache = new Map<string, {
  envName?: string;
  type?: 'conda' | 'venv' | 'pyenv' | 'system';
  version?: string;
}>();

function getPythonInterpreterInfo(pythonPath: string): {
  envName?: string;
  type?: 'conda' | 'venv' | 'pyenv' | 'system';
  version?: string;
} {
  const cached = pythonInterpreterInfoCache.get(pythonPath);
  if (cached) return cached;

  const info = {
    ...detectPythonEnvironment(pythonPath),
    version: detectPythonVersion(pythonPath),
  };
  pythonInterpreterInfoCache.set(pythonPath, info);
  return info;
}

async function getPythonInterpreterInfoAsync(pythonPath: string): Promise<{
  envName?: string;
  type?: 'conda' | 'venv' | 'pyenv' | 'system';
  version?: string;
}> {
  const cached = pythonInterpreterInfoCache.get(pythonPath);
  if (cached) return cached;

  const info = {
    ...detectPythonEnvironment(pythonPath),
    version: await detectPythonVersionAsync(pythonPath),
  };
  pythonInterpreterInfoCache.set(pythonPath, info);
  return info;
}

function detectPythonEnvironment(pythonPath: string): {
  envName?: string;
  type?: 'conda' | 'venv' | 'pyenv' | 'system';
} {
  const normalized = path.normalize(pythonPath);
  const binDir = path.dirname(normalized);
  const binName = path.basename(binDir);
  if (binName !== 'bin' && binName !== 'Scripts') {
    return { type: 'system' };
  }

  const envRoot = path.dirname(binDir);
  const envName = path.basename(envRoot);
  const parentName = path.basename(path.dirname(envRoot));
  const grandParent = path.basename(path.dirname(path.dirname(envRoot)));

  if (parentName === 'envs') {
    return { envName, type: 'conda' };
  }
  if (grandParent === '.pyenv' && parentName === 'versions') {
    return { envName, type: 'pyenv' };
  }
  if (['.venv', 'venv', 'env'].includes(envName) || parentName === '.virtualenvs') {
    return { envName, type: 'venv' };
  }
  if (['miniconda3', 'anaconda3', 'miniforge3', 'mambaforge', 'micromamba'].includes(envName)) {
    return { envName: 'base', type: 'conda' };
  }

  return { envName, type: 'system' };
}

function detectPythonVersion(pythonPath: string): string | undefined {
  try {
    const result = cp.spawnSync(
      pythonPath,
      ['-c', 'import sys; print(".".join(str(p) for p in sys.version_info[:3]))'],
      {
        encoding: 'utf-8',
        timeout: 2500,
        windowsHide: true,
      },
    );
    if (result.status !== 0 || !result.stdout) return undefined;
    const version = result.stdout.trim().split(/\r?\n/).pop()?.trim();
    return version || undefined;
  } catch {
    return undefined;
  }
}

async function detectPythonVersionAsync(pythonPath: string): Promise<string | undefined> {
  const result = await execFileQuiet(
    pythonPath,
    ['-c', 'import sys; print(".".join(str(p) for p in sys.version_info[:3]))'],
    2500,
  );
  if (!result || !result.stdout) return undefined;
  const version = result.stdout.trim().split(/\r?\n/).pop()?.trim();
  return version || undefined;
}

function queryJupyterKernelspecs(): JupyterKernelSpecEntry[] {
  try {
    const result = cp.spawnSync('jupyter', ['kernelspec', 'list', '--json'], {
      encoding: 'utf-8',
      timeout: 5000,
      windowsHide: true,
    });
    if (result.status !== 0 || !result.stdout) return [];

    const parsed = JSON.parse(result.stdout) as {
      kernelspecs?: Record<string, { resource_dir?: string; spec?: JupyterKernelSpec }>;
    };

    return Object.entries(parsed.kernelspecs ?? {}).map(([name, entry]) => ({
      name,
      resourceDir: entry.resource_dir,
      spec: entry.spec ?? {},
    }));
  } catch {
    return [];
  }
}

async function queryJupyterKernelspecsAsync(): Promise<JupyterKernelSpecEntry[]> {
  const result = await execFileQuiet('jupyter', ['kernelspec', 'list', '--json'], 5000);
  if (!result || !result.stdout) return [];

  try {
    const parsed = JSON.parse(result.stdout) as {
      kernelspecs?: Record<string, { resource_dir?: string; spec?: JupyterKernelSpec }>;
    };

    return Object.entries(parsed.kernelspecs ?? {}).map(([name, entry]) => ({
      name,
      resourceDir: entry.resource_dir,
      spec: entry.spec ?? {},
    }));
  } catch {
    return [];
  }
}

function resolvePythonPath(argv0?: string): string {
  if (!argv0) return 'python3';
  if (argv0 === 'python' || argv0 === 'python3') return argv0;
  return pathExists(argv0) ? argv0 : argv0;
}

function resolveRScriptPath(argv0?: string): string | undefined {
  if (!argv0) return undefined;
  if (argv0 === 'Rscript') return argv0;

  const base = path.basename(argv0);
  if (base === 'Rscript') return pathExists(argv0) ? argv0 : undefined;

  if (argv0 === 'R' || base === 'R') {
    const sibling = path.join(path.dirname(argv0), 'Rscript');
    return pathExists(sibling) ? sibling : undefined;
  }

  return undefined;
}

function pathExists(candidate: string): boolean {
  try {
    return fs.existsSync(candidate);
  } catch {
    return false;
  }
}

async function pathExistsAsync(candidate: string): Promise<boolean> {
  try {
    await fs.promises.access(candidate);
    return true;
  } catch {
    return false;
  }
}

function execFileQuiet(
  command: string,
  args: string[],
  timeout: number,
): Promise<{ stdout: string; stderr: string } | null> {
  return new Promise((resolve) => {
    cp.execFile(command, args, {
      encoding: 'utf-8',
      timeout,
      windowsHide: true,
    }, (error, stdout, stderr) => {
      if (error) {
        resolve(null);
        return;
      }
      resolve({ stdout: String(stdout ?? ''), stderr: String(stderr ?? '') });
    });
  });
}

function formatPythonLabel(
  pythonPath: string,
  info: { envName?: string; version?: string },
): string {
  if (info.version && info.envName) return `Python ${info.version} (${info.envName})`;
  if (info.version) return `Python ${info.version}`;
  if (info.envName) return `Python (${info.envName})`;

  const base = path.basename(pythonPath) || pythonPath;
  return base === pythonPath ? `Python (${pythonPath})` : `Python (${base})`;
}

function formatPythonDescription(info: {
  type?: 'conda' | 'venv' | 'pyenv' | 'system';
}): string {
  switch (info.type) {
    case 'conda':
      return 'Conda environment';
    case 'venv':
      return 'Virtual environment';
    case 'pyenv':
      return 'pyenv environment';
    default:
      return 'Python environment';
  }
}
