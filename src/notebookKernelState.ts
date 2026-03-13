export type SelectedRNotebookKernel = {
  controllerId: string;
  displayName: string;
  kernelspecName?: string;
  rPath: string;
};

const rNotebookSelections = new Map<string, SelectedRNotebookKernel>();

export function rememberRNotebookKernel(
  docUri: string,
  selection: SelectedRNotebookKernel,
): void {
  rNotebookSelections.set(docUri, selection);
}

export function getRememberedRNotebookKernel(
  docUri: string,
): SelectedRNotebookKernel | undefined {
  return rNotebookSelections.get(docUri);
}

export function forgetRNotebookKernel(docUri: string): void {
  rNotebookSelections.delete(docUri);
}
