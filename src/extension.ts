import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { SuperDocEditorProvider } from './superdocEditorProvider';
import { BLANK_DOCX_BASE64 } from './blank-docx';

function debug(message: string) {
    console.log('[SuperDoc - Activator]', message);
}

export function activate(context: vscode.ExtensionContext) {
    try {
        debug('Beginning activation...');

        // Register custom editor provider
        const provider = new SuperDocEditorProvider(context);
        const providerRegistration = vscode.window.registerCustomEditorProvider(
            'superdoc.docxEditor',
            provider,
            {
                webviewOptions: {
                    retainContextWhenHidden: true,
                },
                supportsMultipleEditorsPerDocument: false,
            }
        );

        // Register command to open with SuperDoc
        const openCommand = vscode.commands.registerCommand('superdoc.openWithSuperdoc', (uri: vscode.Uri) => {
            vscode.commands.executeCommand('vscode.openWith', uri, 'superdoc.docxEditor');
        });

        // Register URI handler for creating documents
        const uriHandler = vscode.window.registerUriHandler({
            async handleUri(uri: vscode.Uri) {
                debug(`URI handler received: ${uri.toString()}`);

                if (uri.path === '/create') {
                    const params = new URLSearchParams(uri.query);
                    const filePath = params.get('path');

                    if (!filePath) {
                        vscode.window.showErrorMessage('SuperDoc: Missing "path" parameter in URI');
                        return;
                    }

                    try {
                        // Resolve path relative to workspace
                        let absolutePath: string;
                        if (path.isAbsolute(filePath)) {
                            absolutePath = filePath;
                        } else {
                            const workspaceFolder = vscode.workspace.workspaceFolders?.[0];
                            if (!workspaceFolder) {
                                vscode.window.showErrorMessage('SuperDoc: No workspace folder open');
                                return;
                            }
                            absolutePath = path.join(workspaceFolder.uri.fsPath, filePath);
                        }

                        // Ensure .docx extension
                        if (!absolutePath.endsWith('.docx')) {
                            absolutePath += '.docx';
                        }

                        // Ensure parent directory exists
                        const parentDir = path.dirname(absolutePath);
                        if (!fs.existsSync(parentDir)) {
                            fs.mkdirSync(parentDir, { recursive: true });
                        }

                        // Check if file already exists
                        if (fs.existsSync(absolutePath)) {
                            const overwrite = await vscode.window.showWarningMessage(
                                `File "${path.basename(absolutePath)}" already exists. Overwrite?`,
                                'Yes', 'No'
                            );
                            if (overwrite !== 'Yes') {
                                return;
                            }
                        }

                        // Decode base64 and write file
                        const buffer = Buffer.from(BLANK_DOCX_BASE64, 'base64');
                        fs.writeFileSync(absolutePath, buffer);
                        debug(`Created blank document: ${absolutePath}`);

                        // Open the document with SuperDoc
                        const docUri = vscode.Uri.file(absolutePath);
                        await vscode.commands.executeCommand('vscode.openWith', docUri, 'superdoc.docxEditor');

                        vscode.window.showInformationMessage(`Created: ${path.basename(absolutePath)}`);
                    } catch (error) {
                        debug(`Failed to create document: ${error}`);
                        vscode.window.showErrorMessage(`SuperDoc: Failed to create document - ${error}`);
                    }
                }
            }
        });

        context.subscriptions.push(providerRegistration, openCommand, uriHandler);
        debug('Registered extension with URI handler');
        debug('Activation complete');
    } catch (error) {
        debug(`Activation error - ${error}`);
        vscode.window.showErrorMessage(`Failed to activate SuperDoc: ${error}`);
    }
}

export function deactivate() {
    debug('Extension deactivated');
}