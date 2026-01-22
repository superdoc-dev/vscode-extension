import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as https from 'https';
import * as http from 'http';

// Debug logging - set to false to disable verbose logs
const DEBUG_ENABLED = true;

function debug(message: string) {
  if (DEBUG_ENABLED) {
    console.log('[SuperDoc - Provider]', message);
  }
}

// Command folder for Claude API
const SUPERDOC_FOLDER = '.superdoc';

export class SuperDocEditorProvider implements vscode.CustomEditorProvider<SuperDocDocument> {
  public static readonly viewType = 'superdoc.docxEditor';

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<SuperDocDocument>>();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(
    uri: vscode.Uri,
    openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken
  ): Promise<SuperDocDocument> {
    debug(`Opening DOCX file: ${uri.fsPath}`);
    const fileUri = openContext.backupId ? vscode.Uri.parse(openContext.backupId) : uri;
    const data = await vscode.workspace.fs.readFile(fileUri);
    return new SuperDocDocument(uri, data);
  }

  async resolveCustomEditor(
    document: SuperDocDocument,
    webviewPanel: vscode.WebviewPanel,
    _token: vscode.CancellationToken
  ): Promise<void> {
    debug(`Resolving custom editor for: ${document.uri.fsPath}`);

    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [
        vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview'),
      ],
    };

    webviewPanel.webview.html = this.getWebviewContent(webviewPanel.webview);
    this.setupMessageHandler(webviewPanel.webview, document);

    // Send document data when webview is ready
    const readyListener = webviewPanel.webview.onDidReceiveMessage((message) => {
      if (message.type === 'ready') {
        debug(`Sending document to webview, size: ${document.data.length}`);
        webviewPanel.webview.postMessage({
          type: 'update',
          content: { data: Array.from(document.data) },
        });
      }
    });

    // Watch for external file changes
    const fileWatcher = this.setupFileWatcher(document, webviewPanel.webview);

    // Watch for command file (Claude API)
    const commandWatcher = this.setupCommandWatcher(document, webviewPanel.webview);

    webviewPanel.onDidDispose(() => {
      readyListener.dispose();
      fileWatcher.dispose();
      if (commandWatcher) {
        commandWatcher.close();
      }
    });
  }

  private setupFileWatcher(document: SuperDocDocument, webview: vscode.Webview): vscode.FileSystemWatcher {
    const fileDir = vscode.Uri.joinPath(document.uri, '..');
    const fileName = path.basename(document.uri.fsPath);
    const watcher = vscode.workspace.createFileSystemWatcher(
      new vscode.RelativePattern(fileDir, fileName)
    );

    watcher.onDidChange(async (uri) => {
      // Ignore our own saves (within 1 second)
      if (Date.now() - document.lastSaveTime < 1000) {
        debug('Ignoring file change - recent save');
        return;
      }

      debug(`External file change detected: ${uri.fsPath}`);
      await document.reloadFromDisk();
      debug(`Sending reload to webview, size: ${document.data.length} bytes`);
      webview.postMessage({
        type: 'reload',
        content: { data: Array.from(document.data) },
      });
    });

    return watcher;
  }

  /**
   * Get the command file path for a document: .superdoc/{docname}.json
   */
  private getCommandFilePath(documentUri: vscode.Uri): string {
    const docDir = path.dirname(documentUri.fsPath);
    const docName = path.basename(documentUri.fsPath, '.docx');
    return path.join(docDir, SUPERDOC_FOLDER, `${docName}.json`);
  }

  /**
   * Setup watcher for command file (Claude API)
   * Each document watches its own file: .superdoc/{docname}.json
   */
  private setupCommandWatcher(document: SuperDocDocument, webview: vscode.Webview): fs.FSWatcher | null {
    const cmdFilePath = this.getCommandFilePath(document.uri);
    const superdocDir = path.dirname(cmdFilePath);
    const cmdFileName = path.basename(cmdFilePath);
    let processing = false;

    debug(`Setting up command watcher: ${cmdFilePath}`);

    // Ensure .superdoc folder exists
    if (!fs.existsSync(superdocDir)) {
      fs.mkdirSync(superdocDir, { recursive: true });
      debug(`Created folder: ${superdocDir}`);
    }

    const processIfExists = async () => {
      if (processing || !fs.existsSync(cmdFilePath)) return;

      try {
        const content = fs.readFileSync(cmdFilePath, 'utf-8');
        const data = JSON.parse(content);

        // Only process if it's a command (has 'command' field), not a response
        if (!data.command) return;

        processing = true;
        await this.processCommandFile(cmdFilePath, data, webview, document);
      } catch {
        // Ignore parse errors or missing file
      } finally {
        processing = false;
      }
    };

    try {
      const watcher = fs.watch(superdocDir, (eventType, filename) => {
        if (filename === cmdFileName) {
          processIfExists();
        }
      });

      // Check if command file already exists
      processIfExists();

      return watcher;
    } catch (error) {
      debug(`Failed to setup command watcher: ${error}`);
      return null;
    }
  }

  /**
   * Process a command and send to webview
   */
  private async processCommandFile(
    cmdFilePath: string,
    cmd: { id: string; command: string; args?: Record<string, unknown> },
    webview: vscode.Webview,
    document: SuperDocDocument
  ): Promise<void> {
    debug(`Processing command: ${cmd.command} (id: ${cmd.id})`);

    // Store the file path for response writing
    this.pendingCommands.set(cmd.id, cmdFilePath);

    let args = cmd.args || {};

    // Special handling for insertImage - convert URL/path to base64
    if (cmd.command === 'insertImage' && args.src) {
      try {
        args = { ...args };
        const src = args.src as string;

        if (!src.startsWith('data:')) {
          debug(`Converting image source to base64: ${src.substring(0, 100)}...`);
          const docDir = path.dirname(document.uri.fsPath);
          args.src = await this.convertImageToBase64(src, docDir);
          debug(`Image converted, base64 length: ${(args.src as string).length}`);
        }
      } catch (error) {
        // Write error response directly
        this.writeCommandResponse(cmd.id, {
          id: cmd.id,
          success: false,
          error: `Failed to load image: ${error}`
        });
        return;
      }
    }

    webview.postMessage({
      type: 'executeCommand',
      id: cmd.id,
      command: cmd.command,
      args
    });
  }

  /**
   * Convert image URL or file path to base64 data URI
   */
  private async convertImageToBase64(src: string, docDir: string): Promise<string> {
    // Check if it's a URL
    if (src.startsWith('http://') || src.startsWith('https://')) {
      return this.fetchImageAsBase64(src);
    }

    // Otherwise treat as file path
    let filePath = src;
    if (!path.isAbsolute(src)) {
      filePath = path.join(docDir, src);
    }

    if (!fs.existsSync(filePath)) {
      throw new Error(`Image file not found: ${filePath}`);
    }

    const buffer = fs.readFileSync(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const mimeType = this.getMimeType(ext);
    return `data:${mimeType};base64,${buffer.toString('base64')}`;
  }

  /**
   * Fetch image from URL and convert to base64
   */
  private fetchImageAsBase64(url: string): Promise<string> {
    return new Promise((resolve, reject) => {
      const protocol = url.startsWith('https://') ? https : http;

      protocol.get(url, { headers: { 'User-Agent': 'VSCode-SuperDoc/1.0' } }, (response) => {
        // Handle redirects
        if (response.statusCode === 301 || response.statusCode === 302) {
          const redirectUrl = response.headers.location;
          if (redirectUrl) {
            this.fetchImageAsBase64(redirectUrl).then(resolve).catch(reject);
            return;
          }
        }

        if (response.statusCode !== 200) {
          reject(new Error(`HTTP ${response.statusCode}: Failed to fetch image`));
          return;
        }

        const contentType = response.headers['content-type'] || 'image/png';
        const chunks: Buffer[] = [];

        response.on('data', (chunk) => chunks.push(chunk));
        response.on('end', () => {
          const buffer = Buffer.concat(chunks);
          resolve(`data:${contentType};base64,${buffer.toString('base64')}`);
        });
        response.on('error', reject);
      }).on('error', reject);
    });
  }

  /**
   * Get MIME type from file extension
   */
  private getMimeType(ext: string): string {
    const mimeTypes: Record<string, string> = {
      '.png': 'image/png',
      '.jpg': 'image/jpeg',
      '.jpeg': 'image/jpeg',
      '.gif': 'image/gif',
      '.webp': 'image/webp',
      '.svg': 'image/svg+xml',
      '.bmp': 'image/bmp',
      '.ico': 'image/x-icon'
    };
    return mimeTypes[ext] || 'image/png';
  }

  // Track pending commands to know where to write responses
  private pendingCommands = new Map<string, string>();

  /**
   * Write response by overwriting the command file
   */
  private writeCommandResponse(cmdId: string, result: { id: string; success: boolean; result?: unknown; error?: string }): void {
    const cmdFilePath = this.pendingCommands.get(cmdId);
    if (!cmdFilePath) {
      debug(`No pending command found for id: ${cmdId}`);
      return;
    }

    debug(`Writing response: ${result.id} success=${result.success}`);
    fs.writeFileSync(cmdFilePath, JSON.stringify(result, null, 2));
    this.pendingCommands.delete(cmdId);
  }

  private getWebviewContent(webview: vscode.Webview): string {
    const scriptUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview', 'main.js')
    );
    const styleUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview', 'style.css')
    );
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
        <div id="superdoc"></div>
        <script type="module" nonce="${nonce}" src="${scriptUri}"></script>
      </body>
      </html>`;
  }

  private setupMessageHandler(webview: vscode.Webview, document: SuperDocDocument): void {
    webview.onDidReceiveMessage(async (message) => {
      switch (message.type) {
        case 'update': {
          debug(`Received update from webview, size: ${message.content?.length}`);
          document.update(new Uint8Array(message.content));
          await document.save();
          debug(`Document saved: ${document.uri.fsPath}`);
          break;
        }
        case 'debug':
          debug(message.message);
          break;
        case 'commandResult': {
          // Overwrite command file with response
          this.writeCommandResponse(message.id, {
            id: message.id,
            success: message.success,
            result: message.result,
            error: message.error
          });
          break;
        }
      }
    });
  }

  // Required by CustomEditorProvider interface
  async saveCustomDocument(document: SuperDocDocument): Promise<void> {
    await document.save();
  }

  async saveCustomDocumentAs(document: SuperDocDocument, destination: vscode.Uri): Promise<void> {
    await vscode.workspace.fs.writeFile(destination, document.data);
  }

  async revertCustomDocument(document: SuperDocDocument): Promise<void> {
    await document.reloadFromDisk();
  }

  async backupCustomDocument(
    document: SuperDocDocument,
    context: vscode.CustomDocumentBackupContext
  ): Promise<vscode.CustomDocumentBackup> {
    await vscode.workspace.fs.writeFile(context.destination, document.data);
    return {
      id: context.destination.toString(),
      delete: async () => {
        try { await vscode.workspace.fs.delete(context.destination); } catch {}
      },
    };
  }
}

class SuperDocDocument implements vscode.CustomDocument {
  private _data: Uint8Array;
  private _lastSaveTime = 0;

  constructor(public readonly uri: vscode.Uri, initialData: Uint8Array) {
    this._data = initialData;
  }

  get data(): Uint8Array {
    return this._data;
  }

  get lastSaveTime(): number {
    return this._lastSaveTime;
  }

  update(newData: Uint8Array): void {
    this._data = newData;
  }

  async save(): Promise<void> {
    this._lastSaveTime = Date.now();
    await vscode.workspace.fs.writeFile(this.uri, this._data);
  }

  async reloadFromDisk(): Promise<void> {
    this._data = await vscode.workspace.fs.readFile(this.uri);
  }

  dispose(): void {}
}

function getNonce(): string {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  return Array.from({ length: 32 }, () => chars[Math.floor(Math.random() * chars.length)]).join('');
}
