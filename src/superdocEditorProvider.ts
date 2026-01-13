import * as vscode from 'vscode';
import * as path from 'path';

/**
 * Custom editor provider for .docx files using SuperDoc
 */
export class SuperDocEditorProvider implements vscode.CustomEditorProvider<SuperDocDocument> {
  public static readonly viewType = 'superdoc.docxEditor';

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<SuperDocDocument>>();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  constructor(private readonly context: vscode.ExtensionContext) {}

  /**
   * Called when a .docx file is opened
   */
  async openCustomDocument(
    uri: vscode.Uri,
    openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken
  ): Promise<SuperDocDocument> {
    console.log('📁 Opening DOCX file:', uri.fsPath);
    vscode.window.showInformationMessage(`Opening ${uri.fsPath} with SuperDoc`);
    const document = await SuperDocDocument.create(uri, openContext.backupId);
    
    // Listen for document changes to mark as dirty
    const listeners: vscode.Disposable[] = [];
    listeners.push(
      document.onDidChange((e) => {
        this._onDidChangeCustomDocument.fire({
          document,
          undo: e.undo || (() => {}),
          redo: e.redo || (() => {}),
        });
      })
    );

    document.onDidDispose(() => {
      listeners.forEach((l) => l.dispose());
    });

    return document;
  }

  /**
   * Called when we need to display the document in a webview
   */
  async resolveCustomEditor(
    document: SuperDocDocument,
    webviewPanel: vscode.WebviewPanel,
    _token: vscode.CancellationToken
  ): Promise<void> {
    console.log('🌐 Resolving custom editor for:', document.uri.fsPath);
    vscode.window.showInformationMessage('Creating SuperDoc webview...');
    // Setup webview options
    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [
        vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview'),
        vscode.Uri.joinPath(this.context.extensionUri, 'webview'),
        vscode.Uri.joinPath(this.context.extensionUri, 'node_modules'),
      ],
    };

    // Set the webview content
    console.log('📝 Setting webview HTML content');
    webviewPanel.webview.html = this.getWebviewContent(webviewPanel.webview);

    // Handle messages from the webview
    this._setupMessageHandler(webviewPanel.webview, document);

    // Send initial document data once webview is ready
    const sendInitialData = () => {
      const fileData = document.documentData;
      console.log('📤 Sending initial data to webview, size:', fileData.length);
      webviewPanel.webview.postMessage({
        type: 'update',
        content: { data: Array.from(fileData) }, // Convert Uint8Array to regular array for transfer
      });
    };

    // Wait for webview to signal it's ready
    const messageListener = webviewPanel.webview.onDidReceiveMessage((message) => {
      console.log('📨 Message from webview:', message.type);
      if (message.type === 'ready') {
        console.log('✅ Webview is ready, sending initial data');
        sendInitialData();
      }
    });

    // Watch for external file changes (e.g., from LLM or other processes)
    const fileDir = vscode.Uri.joinPath(document.uri, '..');
    const fileName = path.basename(document.uri.fsPath);
    const fileWatcher = vscode.workspace.createFileSystemWatcher(
      new vscode.RelativePattern(fileDir, fileName)
    );

    const SAVE_DEBOUNCE_MS = 1000; // Ignore changes within 1s of our own save

    fileWatcher.onDidChange(async (uri) => {
      if (uri.fsPath !== document.uri.fsPath) {
        return;
      }

      // Ignore if we just saved (to avoid reloading our own changes)
      const timeSinceSave = Date.now() - document.lastSaveTime;
      if (timeSinceSave < SAVE_DEBOUNCE_MS) {
        console.log('🔄 Ignoring file change - recent save detected');
        return;
      }

      console.log('📂 External file change detected, reloading:', uri.fsPath);
      try {
        await document.reloadFromDisk();
      } catch (error) {
        console.error('❌ Failed to reload document:', error);
      }
    });

    // Listen for document content changes (from external reload) and push to webview
    const contentChangeListener = document.onDidChangeContent((newData) => {
      console.log('📤 Pushing reloaded content to webview, size:', newData.length);
      webviewPanel.webview.postMessage({
        type: 'update',
        content: { data: Array.from(newData) },
      });
    });

    webviewPanel.onDidDispose(() => {
      messageListener.dispose();
      fileWatcher.dispose();
      contentChangeListener.dispose();
    });
  }

  /**
   * Generate HTML content for the webview
   */
  private getWebviewContent(webview: vscode.Webview): string {
    const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(
      this.context.extensionUri, 'dist', 'webview', 'main.js'
    ));
    
    const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(
      this.context.extensionUri, 'dist', 'webview', 'style.css'
    ));

    const nonce = getNonce();

    return `<!DOCTYPE html>
      <html lang="en">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <meta http-equiv="Content-Security-Policy" content="default-src 'none'; font-src ${webview.cspSource} https: data:; img-src ${webview.cspSource} https: data: blob:; script-src 'nonce-${nonce}' ${webview.cspSource}; style-src ${webview.cspSource} 'unsafe-inline';">
          <link href="${styleUri}" rel="stylesheet">
          <title>SuperDoc Editor</title>
      </head>
      <body>
          <div id="superdoc-toolbar"></div>
          <div id="superdoc">
            <div style="color: white; padding: 20px; background: #333;">
              Initial HTML loaded... waiting for script
            </div>
          </div>
          <script type="module" nonce="${nonce}" src="${scriptUri}"></script>
      </body>
      </html>`;
  }

  /**
   * Setup message handler for webview <-> extension communication
   */
  private _setupMessageHandler(
    webview: vscode.Webview,
    document: SuperDocDocument
  ): void {
    webview.onDidReceiveMessage(async (message) => {
      switch (message.type) {
        case 'update':
          // Document was edited in SuperDoc
          console.log('📥 Received update from webview, size:', message.content?.length);
          const newContent = new Uint8Array(message.content);
          document.update(newContent);
          // Auto-save to disk
          await document.save();
          console.log('💾 Document saved to disk:', document.uri.fsPath);
          break;

        case 'save':
          // Explicit save request
          await document.save();
          webview.postMessage({ type: 'saved' });
          break;

        case 'error':
          vscode.window.showErrorMessage(`SuperDoc Error: ${message.message}`);
          break;

        case 'info':
          vscode.window.showInformationMessage(message.message);
          break;

        case 'debug':
          console.log('[Webview]', message.message);
          break;
      }
    });
  }

  /**
   * Called when the user saves the document
   */
  async saveCustomDocument(
    document: SuperDocDocument,
    cancellation: vscode.CancellationToken
  ): Promise<void> {
    await document.save(cancellation);
  }

  /**
   * Called when the user does "Save As"
   */
  async saveCustomDocumentAs(
    document: SuperDocDocument,
    destination: vscode.Uri,
    cancellation: vscode.CancellationToken
  ): Promise<void> {
    await document.saveAs(destination, cancellation);
  }

  /**
   * Called when VS Code wants to revert the document
   */
  async revertCustomDocument(
    document: SuperDocDocument,
    cancellation: vscode.CancellationToken
  ): Promise<void> {
    await document.revert(cancellation);
  }

  /**
   * Called when VS Code wants to create a backup
   */
  async backupCustomDocument(
    document: SuperDocDocument,
    context: vscode.CustomDocumentBackupContext,
    cancellation: vscode.CancellationToken
  ): Promise<vscode.CustomDocumentBackup> {
    return document.backup(context.destination, cancellation);
  }
}

/**
 * Represents a .docx document managed by SuperDoc
 */
export class SuperDocDocument implements vscode.CustomDocument {
  private _documentData: Uint8Array;
  private _savedData: Uint8Array;
  private _lastSaveTime: number = 0;
  private readonly _onDidChange = new vscode.EventEmitter<{
    readonly undo?: () => void;
    readonly redo?: () => void;
  }>();
  public readonly onDidChange = this._onDidChange.event;

  private readonly _onDidDispose = new vscode.EventEmitter<void>();
  public readonly onDidDispose = this._onDidDispose.event;

  private readonly _onDidChangeContent = new vscode.EventEmitter<Uint8Array>();
  public readonly onDidChangeContent = this._onDidChangeContent.event;

  private constructor(
    public readonly uri: vscode.Uri,
    initialData: Uint8Array
  ) {
    this._documentData = initialData;
    this._savedData = initialData;
  }

  get lastSaveTime(): number {
    return this._lastSaveTime;
  }

  /**
   * Create a new SuperDocDocument
   */
  static async create(
    uri: vscode.Uri,
    backupId?: string
  ): Promise<SuperDocDocument> {
    // If we have a backup, read that instead
    const fileUri = backupId ? vscode.Uri.parse(backupId) : uri;
    const data = await vscode.workspace.fs.readFile(fileUri);
    return new SuperDocDocument(uri, data);
  }

  /**
   * Get the current document data
   */
  get documentData(): Uint8Array {
    return this._documentData;
  }

  /**
   * Update the document with new content from SuperDoc
   */
  update(newData: Uint8Array): void {
    const oldData = this._documentData;
    this._documentData = newData;

    // Fire change event with undo/redo support
    this._onDidChange.fire({
      undo: () => {
        this._documentData = oldData;
      },
      redo: () => {
        this._documentData = newData;
      },
    });
  }

  /**
   * Save the document to disk
   */
  async save(cancellation?: vscode.CancellationToken): Promise<void> {
    this._lastSaveTime = Date.now();
    await this.saveAs(this.uri, cancellation);
    this._savedData = this._documentData;
  }

  /**
   * Save the document to a different location
   */
  async saveAs(
    targetUri: vscode.Uri,
    _cancellation?: vscode.CancellationToken
  ): Promise<void> {
    await vscode.workspace.fs.writeFile(targetUri, this._documentData);
  }

  /**
   * Revert to the last saved version
   */
  async revert(_cancellation?: vscode.CancellationToken): Promise<void> {
    const data = await vscode.workspace.fs.readFile(this.uri);
    this._documentData = data;
    this._savedData = data;
  }

  /**
   * Reload from disk (for external changes) and notify listeners
   */
  async reloadFromDisk(): Promise<void> {
    const data = await vscode.workspace.fs.readFile(this.uri);
    this._documentData = data;
    this._savedData = data;
    this._onDidChangeContent.fire(data);
  }

  /**
   * Create a backup of the document
   */
  async backup(
    destination: vscode.Uri,
    _cancellation?: vscode.CancellationToken
  ): Promise<vscode.CustomDocumentBackup> {
    await vscode.workspace.fs.writeFile(destination, this._documentData);
    return {
      id: destination.toString(),
      delete: async () => {
        try {
          await vscode.workspace.fs.delete(destination);
        } catch {
          // Ignore - backup may not exist
        }
      },
    };
  }

  dispose(): void {
    this._onDidDispose.fire();
    this._onDidChange.dispose();
    this._onDidChangeContent.dispose();
    this._onDidDispose.dispose();
  }
}

function getNonce() {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}
