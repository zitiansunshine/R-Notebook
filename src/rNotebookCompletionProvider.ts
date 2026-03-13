import * as vscode from 'vscode';

import { NOTEBOOK_TYPE, getRConfigValue } from './extensionIds';
import { getRememberedRNotebookKernel } from './notebookKernelState';
import { getOrCreateSession } from './rSessionManager';

function resolveNotebook(document: vscode.TextDocument): vscode.NotebookDocument | undefined {
  return vscode.workspace.notebookDocuments.find((notebook) =>
    notebook.notebookType === NOTEBOOK_TYPE
    && notebook.getCells().some((cell) => cell.document.uri.toString() === document.uri.toString()),
  );
}

function memberCompletionRange(
  document: vscode.TextDocument,
  position: vscode.Position,
): vscode.Range | undefined {
  const linePrefix = document.lineAt(position.line).text.slice(0, position.character);
  const match = /(?:\$|@)([A-Za-z0-9._]*)$/.exec(linePrefix);
  if (!match) return undefined;
  return new vscode.Range(position.line, position.character - match[1].length, position.line, position.character);
}

function memberName(completion: string): string {
  const splitIndex = Math.max(completion.lastIndexOf('$'), completion.lastIndexOf('@'));
  return splitIndex >= 0 ? completion.slice(splitIndex + 1) : completion;
}

export function registerRNotebookCompletionProvider(): vscode.Disposable {
  const selector: vscode.DocumentSelector = [
    { language: 'r', scheme: 'vscode-notebook-cell' },
    { language: 'R', scheme: 'vscode-notebook-cell' },
  ];

  return vscode.languages.registerCompletionItemProvider(selector, {
    async provideCompletionItems(document, position, token, context) {
      if (
        context.triggerKind === vscode.CompletionTriggerKind.TriggerCharacter
        && context.triggerCharacter !== '$'
        && context.triggerCharacter !== '@'
      ) {
        return undefined;
      }

      const range = memberCompletionRange(document, position);
      if (!range) return undefined;

      const notebook = resolveNotebook(document);
      if (!notebook) return undefined;
      const selectedKernel = getRememberedRNotebookKernel(notebook.uri.toString());

      const session = getOrCreateSession(
        notebook.uri.toString(),
        selectedKernel?.rPath ?? getRConfigValue('rPath', 'Rscript'),
        getRConfigValue('execTimeoutMs', 0) ?? 0,
      );

      if (session.isBusy()) return undefined;

      let completions: string[];
      try {
        completions = await session.complete(
          `nbcmp-${position.line}-${position.character}`,
          document.getText(),
          document.offsetAt(position),
        );
      } catch {
        return undefined;
      }

      if (token.isCancellationRequested || completions.length === 0) return undefined;

      const seen = new Set<string>();
      const items: vscode.CompletionItem[] = [];
      for (const completion of completions) {
        const label = memberName(completion);
        if (!label || seen.has(label)) continue;
        seen.add(label);

        const item = new vscode.CompletionItem(label, vscode.CompletionItemKind.Field);
        item.detail = completion;
        item.insertText = label;
        item.filterText = label;
        item.range = range;
        items.push(item);
      }

      return items.length > 0 ? new vscode.CompletionList(items, false) : undefined;
    },
  }, '$', '@');
}
