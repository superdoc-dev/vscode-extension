// SuperDoc integration for VS Code webview

// Import SuperDoc styles as text and inject
import superdocCss from 'superdoc/style.css';

// Debug function to write messages to a debug div and VS Code
let vscode = null; // Will be set after acquireVsCodeApi()

function debug(message) {
    console.log(message);

    // Post to VS Code for logging in output panel
    if (vscode) {
        vscode.postMessage({ type: 'debug', message });
    }

    if (document && document.body) {
        // Create or update debug div without destroying other elements
        let debugDiv = document.getElementById('debug-output');
        if (!debugDiv) {
            debugDiv = document.createElement('div');
            debugDiv.id = 'debug-output';
            debugDiv.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; z-index: 9999; color: white; padding: 10px; background: #333; font-family: monospace; font-size: 12px; border-bottom: 1px solid #555;';
            document.body.appendChild(debugDiv);
        }
        debugDiv.textContent = message;
    }
}

debug('🚀 Webview main.js loading...');

// Test if DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeWebview);
} else {
    initializeWebview();
}

function initializeWebview() {
    debug('📱 DOM ready, initializing webview...');
    
    // Inject SuperDoc CSS
    if (superdocCss) {
        const style = document.createElement('style');
        style.textContent = superdocCss;
        document.head.appendChild(style);
        debug('📎 SuperDoc CSS injected');
    }
}

vscode = acquireVsCodeApi();
debug('📱 VS Code API acquired');

// Import SuperDoc - will be bundled by esbuild
import { SuperDoc } from 'superdoc';
debug('📚 SuperDoc imported');

let editor = null;
let currentFileData = null;
let saveTimeout = null;
let isInitialLoad = true;

// Configuration
const AUTO_SAVE_DELAY = 1000; // milliseconds

// Initialize editor with file data
function initializeEditor(fileArrayBuffer) {
    debug('🎯 Initializing editor with file buffer');

    try {
        // Convert ArrayBuffer to File object
        const file = new File([fileArrayBuffer], 'document.docx', {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        debug(`📄 File created: ${file.name}, ${file.size} bytes`);

        // Clean up previous editor instance
        if (editor) {
            debug('🧹 Cleaning up previous editor instance');
            try {
                if (editor.destroy) {
                    editor.destroy();
                }
            } catch (e) {
                debug(`⚠️ Error destroying editor: ${e.message}`);
            }
            editor = null;
        }

        // Reset state for new editor
        isInitialLoad = true;

        // Check if DOM elements exist
        const superdocElement = document.getElementById('superdoc');
        const toolbarElement = document.getElementById('superdoc-toolbar');

        // Clear existing content from containers
        if (superdocElement) {
            superdocElement.innerHTML = '';
        }
        if (toolbarElement) {
            toolbarElement.innerHTML = '';
        }
        
        if (!superdocElement) {
            debug('❌ #superdoc element not found!');
            return;
        }
        if (!toolbarElement) {
            debug('❌ #superdoc-toolbar element not found!');
            return;
        }
        
        debug('✅ DOM elements found, creating SuperDoc...');

        // Initialize SuperDoc
        debug('🚀 Creating SuperDoc instance...');
        
        try {
            editor = new SuperDoc({
                selector: '#superdoc',
                toolbar: '#superdoc-toolbar',
                document: file,
                documentMode: 'editing',
                pagination: true,
                rulers: true,
                onReady: (event) => {
                    debug('✅ SuperDoc is ready');
                    isInitialLoad = false;
                    setupEditorListeners();
                },
                onEditorCreate: (event) => {
                    debug('✅ Editor is created');
                },
                onError: (error) => {
                    debug(`❌ SuperDoc error: ${error.message || error}`);
                }
            });
            
            debug('✅ SuperDoc constructor completed');
            
            // Set a timeout to check if SuperDoc initializes
            setTimeout(() => {
                if (!isInitialLoad) {
                    debug('⏱️ SuperDoc initialized successfully');
                } else {
                    debug('⏱️ SuperDoc still initializing after 5 seconds...');
                }
            }, 5000);
            
        } catch (constructorError) {
            debug(`❌ SuperDoc constructor failed: ${constructorError.message}`);
        }
        
    } catch (error) {
        debug(`❌ Failed to initialize SuperDoc: ${error.message}`);
    }
}

// Setup editor listeners for content changes (debounced save on update)
function setupEditorListeners() {
    if (!editor?.activeEditor) {
        debug('❌ No editor or activeEditor available');
        return;
    }

    debug('🎧 Setting up editor update listener...');

    // Listen for editor updates - use this for debounced save
    editor.activeEditor.on('update', async ({ editor: editorInstance }) => {
        if (isInitialLoad) return;

        // Log the HTML representation of the editor on each update
        const html = editorInstance.getHTML();
        debug(`📝 Content updated: ${html?.length || 0} chars`);

        // Schedule debounced auto-save
        scheduleAutoSave();
    });

    debug('✅ Editor update listener ready');
}

// Schedule auto-save with debouncing
function scheduleAutoSave() {
    // Clear existing timeout
    if (saveTimeout) {
        clearTimeout(saveTimeout);
    }

    // Schedule new save
    saveTimeout = setTimeout(() => {
        saveDocument();
    }, AUTO_SAVE_DELAY);
}

// Save document back to VS Code
async function saveDocument() {
    if (!editor) {
        debug('❌ No editor available for saving');
        return;
    }

    try {
        debug('💾 Starting document save...');
        
        // Export the document as blob
        const blob = await editor.export({
            format: 'docx'
        });

        if (!blob) {
            debug('❌ Failed to export - no blob returned');
            return;
        }

        debug(`📦 Exported blob size: ${blob.size} bytes`);

        // Convert blob to ArrayBuffer, then to plain array for serialization
        const arrayBuffer = await blob.arrayBuffer();
        const contentArray = Array.from(new Uint8Array(arrayBuffer));

        // Send the updated content back to VS Code
        vscode.postMessage({
            type: 'update',
            content: contentArray
        });

        debug(`✅ Document sent to VS Code (${contentArray.length} bytes)`);
    } catch (error) {
        debug(`❌ Error saving document: ${error.message}`);
    }
}

// Handle messages from VS Code
window.addEventListener('message', async event => {
    const message = event.data;
    debug(`📨 Received message from VS Code: ${message.type}`);
    
    switch (message.type) {
        case 'update':
            // Receive document data from VS Code
            if (message.content && message.content.data) {
                currentFileData = new Uint8Array(message.content.data).buffer;
                debug(`📊 File data received: ${currentFileData.byteLength} bytes`);
                initializeEditor(currentFileData);
            }
            break;
    }
});

// Notify VS Code that the webview is ready
debug('✋ Notifying VS Code that webview is ready');
vscode.postMessage({ type: 'ready' });

// Handle keyboard shortcuts
document.addEventListener('keydown', (event) => {
    // Ctrl/Cmd + S to save
    if ((event.ctrlKey || event.metaKey) && event.key === 's') {
        event.preventDefault();
        saveDocument();
        vscode.postMessage({ type: 'save' });
    }
});