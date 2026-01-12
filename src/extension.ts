import * as vscode from 'vscode';
import { SuperDocEditorProvider } from './superdocEditorProvider';

export function activate(context: vscode.ExtensionContext) {
    try {
        console.log('🚀 SuperDoc extension is now active!');
        vscode.window.showInformationMessage('SuperDoc extension activated!');

        // Register the custom editor provider
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

        // Register the command to open with SuperDoc
        const openCommand = vscode.commands.registerCommand('superdoc.openWithSuperdoc', (uri: vscode.Uri) => {
            vscode.commands.executeCommand('vscode.openWith', uri, 'superdoc.docxEditor');
        });

        context.subscriptions.push(providerRegistration, openCommand);
        console.log('✅ SuperDoc extension registration complete!');
    } catch (error) {
        console.error('❌ Error activating SuperDoc extension:', error);
        vscode.window.showErrorMessage(`Failed to activate SuperDoc: ${error}`);
    }
}

export function deactivate() {
    console.log('SuperDoc extension is now deactivated');
}