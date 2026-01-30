// SuperDoc integration for VS Code webview
/* global document, window, setTimeout, clearTimeout, acquireVsCodeApi, File */

import superdocCss from 'superdoc/style.css';
import { SuperDoc } from 'superdoc';

// =============================================================================
// Configuration & Setup
// =============================================================================

const DEBUG_ENABLED = false;
const AUTO_SAVE_DELAY = 1000;

// Suppress noisy SuperDoc internal logs
const originalConsoleLog = console.log;
console.log = (...args) => {
    if (typeof args[0] === 'string' && args[0].includes('[sd-table-borders]')) return;
    originalConsoleLog.apply(console, args);
};

const vscode = acquireVsCodeApi();

function debug(message) {
    if (DEBUG_ENABLED && vscode) {
        vscode.postMessage({ type: 'debug', message: `[Webview] - ${message}` });
    }
}

// Inject CSS when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injectStyles);
} else {
    injectStyles();
}

function injectStyles() {
    if (superdocCss) {
        const style = document.createElement('style');
        style.textContent = superdocCss;
        document.head.appendChild(style);
    }
}

let editor = null;
let saveTimeout = null;
let isInitialLoad = true;

// Initialize editor with file data
function initializeEditor(fileArrayBuffer) {
    debug('Initializing editor with file buffer');

    try {
        // Convert ArrayBuffer to File object
        const file = new File([fileArrayBuffer], 'document.docx', {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        debug(`File created: ${file.name}, ${file.size} bytes`);

        // Clean up previous editor instance
        if (editor) {
            debug('Destroying previous editor...');
            try {
                if (editor.destroy) {
                    editor.destroy();
                }
            } catch (e) {
                debug(`Error destroying previous editor: ${e.message}`);
            }
            editor = null;
            debug('Destroyed previous editor');
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
        
        if (!superdocElement || !toolbarElement) {
            throw new Error('Required DOM elements not found (#superdoc or #superdoc-toolbar)');
        }
        
        debug('DOM elements found, creating SuperDoc...');
        
        try {
            editor = new SuperDoc({
                selector: '#superdoc',
                toolbar: '#superdoc-toolbar',
                document: file,
                documentMode: 'editing',  // Default to normal editing (commands switch to suggesting mode for tracked changes)
                pagination: true,
                rulers: true,
                user: {
                    name: 'Claude',
                    email: 'claude@anthropic.com'
                },
                onReady: () => {
                    debug('SuperDoc is ready (editing mode)');
                    isInitialLoad = false;
                    setupEditorListeners();
                },
                onEditorCreate: () => {
                    debug('Editor created');
                },
                onError: (error) => {
                    debug(`SuperDoc error: ${error.message || error}`);
                }
            });
            
            debug('SuperDoc init complete');
            
        } catch (constructorError) {
            debug(`SuperDoc constructor failed: ${constructorError.message}`);
        }

    } catch (error) {
        debug(`Failed to initialize SuperDoc: ${error.message}`);
    }
}

// Setup editor listeners for content changes (debounced save on update)
function setupEditorListeners() {
    if (!editor?.activeEditor) {
        debug('No editor or activeEditor available');
        return;
    }

    debug('Setting up editor update listener');

    editor.activeEditor.on('update', async ({ editor: editorInstance }) => {
        if (isInitialLoad) return;

        const html = editorInstance.getHTML();
        debug(`Content updated: ${html?.length || 0} chars`);
        scheduleAutoSave();
    });

    debug('Editor update listener ready');
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
        debug('No editor available for saving');
        return;
    }

    try {
        debug('Starting document save...');

        const blob = await editor.export({ format: 'docx' });
        if (!blob) {
            debug('Failed to export - no blob returned');
            return;
        }

        debug(`Exported blob size: ${blob.size} bytes`);

        const arrayBuffer = await blob.arrayBuffer();
        const contentArray = Array.from(new Uint8Array(arrayBuffer));

        vscode.postMessage({
            type: 'update',
            content: contentArray
        });

        debug(`Document sent to VS Code (${contentArray.length} bytes)`);
    } catch (error) {
        debug(`Error saving document: ${error.message}`);
    }
}

// Handle messages from VS Code
window.addEventListener('message', async event => {
    const message = event.data;
    if (!message?.type) {
        return;
    }

    debug(`Received message: ${message.type}`);

    switch (message.type) {
        case 'update':
        case 'reload':
            if (message.content?.data) {
                const fileBuffer = new Uint8Array(message.content.data).buffer;
                debug(`${message.type}: ${fileBuffer.byteLength} bytes`);
                initializeEditor(fileBuffer);
            }
            break;

        case 'executeCommand':
            debug(`Executing command: ${message.command}`);
            const result = await executeCommand(message.command, message.args || {});
            vscode.postMessage({ type: 'commandResult', ...result });
            break;
    }
});

// =============================================================================
// Command Execution API
// =============================================================================

const COMMANDS = {
    getText: cmdGetText,
    getNodes: cmdGetNodes,
    replaceText: cmdReplaceText,
    insertContent: cmdInsertContent,
    formatText: cmdFormatText,
    insertImage: cmdInsertImage,
    deleteNode: cmdDeleteNode,
    insertTable: cmdInsertTable,
    addComment: cmdAddComment,
    insertTableOfContents: cmdInsertTableOfContents,
    deleteTableOfContents: cmdDeleteTableOfContents,
    undo: cmdUndo,
    redo: cmdRedo
};

async function executeCommand(command, args) {
    const handler = COMMANDS[command];
    if (!handler) {
        return { success: false, error: `Unknown command: ${command}` };
    }
    try {
        return await handler(args);
    } catch (error) {
        debug(`Command error: ${error.message}`);
        return { success: false, error: error.message };
    }
}

// =============================================================================
// Editor Helpers
// =============================================================================

function getActiveEditor() {
    return editor?.activeEditor || null;
}

function requireActiveEditor() {
    const activeEditor = getActiveEditor();
    if (!activeEditor) {
        return { error: { success: false, error: 'No active editor' } };
    }
    return { activeEditor };
}

function getFormattedText(activeEditor) {
    try {
        return activeEditor.getText({ blockSeparator: '\n\n' });
    } catch {
        return activeEditor.state.doc.textContent;
    }
}

function setDocumentMode(mode) {
    editor.setDocumentMode(mode);
    if (mode === 'suggesting') {
        getActiveEditor()?.commands.enableTrackChanges?.();
    }
}

function setAuthorIfProvided(author) {
    if (author?.name) {
        const user = {
            name: author.name,
            email: author.email || `${author.name.toLowerCase().replace(/\s+/g, '.')}@user.local`
        };
        // Set on both SuperDoc wrapper and Editor instance (track changes reads from editor.options.user)
        editor.user = user;
        const activeEditor = getActiveEditor();
        if (activeEditor) {
            activeEditor.options.user = user;
        }
        debug(`Author set to: ${user.name}`);
    }
}

function searchText(activeEditor, text) {
    return activeEditor.commands.search(text, { highlight: false }) || [];
}

function findAnchor(activeEditor, anchor) {
    const matches = searchText(activeEditor, anchor);
    return matches[0] || null;
}

function findMatch(activeEditor, search, occurrence) {
    const matches = searchText(activeEditor, search);
    if (matches.length === 0) {
        return { error: `Text not found: "${search}"` };
    }
    const idx = occurrence ? parseInt(occurrence, 10) - 1 : 0;
    if (idx < 0 || idx >= matches.length) {
        return { error: `Occurrence ${occurrence} not found (only ${matches.length} matches)` };
    }
    return { match: matches[idx], matches };
}

function insertAtPosition(activeEditor, position, content) {
    activeEditor.view.focus();
    activeEditor.commands.setTextSelection({ from: position, to: position });
    activeEditor.commands.insertContent(content);
}

function selectRange(activeEditor, from, to) {
    activeEditor.view.focus();
    activeEditor.commands.setTextSelection({ from, to: to ?? from });
}

function applyScope(activeEditor, scope) {
    activeEditor.view.focus();
    if (scope === 'document') {
        activeEditor.commands.selectAll();
    } else if (scope?.from !== undefined && scope?.to !== undefined) {
        activeEditor.commands.setTextSelection({ from: scope.from, to: scope.to });
    }
}

function cmdGetText({ format } = {}) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const validFormats = ['text', 'html', 'both'];
    const selectedFormat = format || 'both';

    if (!validFormats.includes(selectedFormat)) {
        return { success: false, error: `Invalid format: "${format}". Valid formats: ${validFormats.join(', ')}` };
    }

    const result = {};

    if (selectedFormat === 'text' || selectedFormat === 'both') {
        result.text = getFormattedText(activeEditor);
    }

    if (selectedFormat === 'html' || selectedFormat === 'both') {
        result.html = activeEditor.getHTML?.() || null;
    }

    const charCount = result.text?.length || result.html?.length || 0;
    debug(`getText(${selectedFormat}): ${charCount} chars`);
    return { success: true, result };
}

function cmdGetNodes({ type }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!type) return { success: false, error: 'Node type is required' };

    const validTypes = ['paragraph', 'table', 'tableRow', 'tableCell',
                        'bulletList', 'orderedList', 'listItem', 'image', 'blockquote'];

    if (!validTypes.includes(type)) {
        return { success: false, error: `Invalid type: "${type}". Valid types: ${validTypes.join(', ')}` };
    }

    const nodes = activeEditor.getNodesOfType(type);

    const result = nodes.map((item, index) => {
        const { node, pos } = item;
        const from = pos;
        const to = pos + node.nodeSize;
        const text = node.textContent || '';

        const entry = {
            index,
            type,
            from,
            to,
            text: text.substring(0, 100) + (text.length > 100 ? '...' : ''),
            textLength: text.length,
        };

        // For paragraphs, include numbering marker if present (e.g., "1.", "(A)", "•")
        if (type === 'paragraph') {
            try {
                const dom = activeEditor.view.nodeDOM(pos);
                const marker = dom?.getAttribute?.('data-marker-type');
                if (marker) entry.marker = marker;
            } catch {}
        }

        return entry;
    });

    debug(`getNodes: found ${result.length} ${type} nodes`);
    return { success: true, result: { nodes: result, count: result.length } };
}

async function cmdFormatText({ fontFamily, fontSize, color, highlight, bold, italic, underline, strikethrough, link, lineHeight, indent, spacingBefore, spacingAfter, scope }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const hasFormat = fontFamily || fontSize || color || highlight !== undefined ||
                      bold !== undefined || italic !== undefined ||
                      underline !== undefined || strikethrough !== undefined ||
                      link !== undefined ||
                      lineHeight !== undefined || indent !== undefined ||
                      spacingBefore !== undefined || spacingAfter !== undefined;
    if (!hasFormat) {
        return { success: false, error: 'At least one format option required' };
    }

    const previousMode = editor.documentMode;
    editor.setDocumentMode('editing');
    applyScope(activeEditor, scope);

    const applied = [];

    // Font properties
    if (fontFamily) {
        activeEditor.commands.setFontFamily(fontFamily);
        applied.push(`fontFamily: ${fontFamily}`);
    }

    if (fontSize) {
        activeEditor.commands.setFontSize(fontSize);
        applied.push(`fontSize: ${fontSize}`);
    }

    if (color) {
        activeEditor.commands.setColor(color);
        applied.push(`color: ${color}`);
    }

    // Highlight (background color) - string = set color, false = remove
    if (highlight && highlight !== false) {
        activeEditor.commands.setHighlight(highlight);
        applied.push(`highlight: ${highlight}`);
    } else if (highlight === false) {
        activeEditor.commands.unsetHighlight();
        applied.push('highlight: removed');
    }

    // Text formatting (true = set, false = unset)
    if (bold === true) {
        activeEditor.commands.setBold();
        applied.push('bold: true');
    } else if (bold === false) {
        activeEditor.commands.unsetBold();
        applied.push('bold: false');
    }

    if (italic === true) {
        activeEditor.commands.setItalic();
        applied.push('italic: true');
    } else if (italic === false) {
        activeEditor.commands.unsetItalic();
        applied.push('italic: false');
    }

    if (underline === true) {
        activeEditor.commands.setUnderline();
        applied.push('underline: true');
    } else if (underline === false) {
        activeEditor.commands.unsetUnderline();
        applied.push('underline: false');
    }

    if (strikethrough === true) {
        activeEditor.commands.setStrike();
        applied.push('strikethrough: true');
    } else if (strikethrough === false) {
        activeEditor.commands.unsetStrike();
        applied.push('strikethrough: false');
    }

    // Link - string = set link href, false = remove link
    if (link && link !== false) {
        activeEditor.commands.setLink({ href: link });
        applied.push(`link: ${link}`);
    } else if (link === false) {
        activeEditor.commands.unsetLink();
        applied.push('link: removed');
    }

    // Indentation (in points) — uses setTextIndentation which sets paragraphProperties.indent.left
    if (indent !== undefined) {
        const indentPoints = parseFloat(indent);
        if (!Number.isNaN(indentPoints)) {
            if (indentPoints === 0) {
                activeEditor.commands.unsetTextIndentation();
            } else {
                activeEditor.commands.setTextIndentation(indentPoints);
            }
            applied.push(`indent: ${indent}`);
        }
    }

    // Line height — uses setLineHeight which sets paragraphProperties.spacing.line
    if (lineHeight !== undefined) {
        const lh = parseFloat(lineHeight);
        if (!Number.isNaN(lh)) {
            if (lh === 0) {
                activeEditor.commands.unsetLineHeight();
            } else {
                activeEditor.commands.setLineHeight(lh);
            }
            applied.push(`lineHeight: ${lineHeight}`);
        }
    }

    // Spacing before/after — set via paragraphProperties.spacing (values in twips, 1pt = 20 twips)
    if (spacingBefore !== undefined) {
        const pts = parseFloat(spacingBefore);
        if (!Number.isNaN(pts)) {
            activeEditor.commands.updateAttributes('paragraph', {
                'paragraphProperties.spacing.before': pts * 20,
            });
            applied.push(`spacingBefore: ${spacingBefore}`);
        }
    }
    if (spacingAfter !== undefined) {
        const pts = parseFloat(spacingAfter);
        if (!Number.isNaN(pts)) {
            activeEditor.commands.updateAttributes('paragraph', {
                'paragraphProperties.spacing.after': pts * 20,
            });
            applied.push(`spacingAfter: ${spacingAfter}`);
        }
    }

    // Restore previous mode
    editor.setDocumentMode(previousMode);

    await saveDocument();
    debug(`formatText: applied ${applied.join(', ')}`);
    return { success: true, result: { applied } };
}

async function cmdReplaceText({ search, replacement, occurrence, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!search) return { success: false, error: 'Search text is required' };

    // Search for matches first (before changing document mode)
    const matches = searchText(activeEditor, search);
    if (matches.length === 0) {
        return { success: false, error: `Text not found: "${search}"` };
    }

    // Validate occurrence if specified
    let toReplace;
    if (occurrence != null) {
        const idx = parseInt(occurrence, 10) - 1;
        if (idx < 0 || idx >= matches.length) {
            return { success: false, error: `Occurrence ${occurrence} not found (only ${matches.length} matches)` };
        }
        toReplace = [matches[idx]];
    } else {
        // Reverse to maintain positions when replacing multiple
        toReplace = [...matches].reverse();
    }

    setAuthorIfProvided(author);
    setDocumentMode('suggesting');

    debug(`replaceText: found ${matches.length} matches`);

    // Replace using proper positions from search
    for (const m of toReplace) {
        selectRange(activeEditor, m.from, m.to);
        activeEditor.commands.insertContent(replacement);
    }

    await saveDocument();
    setDocumentMode('editing');
    debug(`replaceText: replaced ${toReplace.length} occurrence(s)`);
    return { success: true, result: { replacedCount: toReplace.length } };
}

async function cmdInsertContent({ content, position, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!content) return { success: false, error: 'Content is required' };

    const anchor = position?.after || position?.before;
    const insertAfter = Boolean(position?.after);
    const isEmptyDoc = activeEditor.state.doc.textContent.trim().length === 0;

    // Validate before changing mode
    if (!anchor && !isEmptyDoc) {
        return { success: false, error: 'Position anchor required: use "after" or "before" with existing text' };
    }

    // Find anchor before changing mode (if needed)
    let insertPos = 1;
    if (!isEmptyDoc) {
        const match = findAnchor(activeEditor, anchor);
        if (!match) {
            return { success: false, error: `Anchor text not found: "${anchor}"` };
        }
        insertPos = insertAfter ? match.to : match.from;
    }

    setAuthorIfProvided(author);
    setDocumentMode('suggesting');

    insertAtPosition(activeEditor, insertPos, content);
    debug(isEmptyDoc ? 'insertContent: empty document' : `insertContent: ${insertAfter ? 'after' : 'before'} "${anchor}"`);

    await saveDocument();
    setDocumentMode('editing');
    return { success: true };
}

async function cmdInsertImage({ src, alt, width, position }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!src) return { success: false, error: 'Image src is required (URL or base64 data URI)' };

    const anchor = position?.after || position?.before;
    if (!anchor) {
        return { success: false, error: 'Position anchor (after/before) is required' };
    }

    // Use edit mode for images (not track changes)
    setDocumentMode('editing');

    const match = findAnchor(activeEditor, anchor);
    if (!match) {
        return { success: false, error: `Anchor text not found: "${anchor}"` };
    }

    const imageAttrs = { src };
    if (alt) imageAttrs.alt = alt;
    if (width) imageAttrs.size = { width };

    try {
        const insertPos = position?.after ? match.to : match.from;
        insertAtPosition(activeEditor, insertPos, { type: 'image', attrs: imageAttrs });
    } catch (e) {
        return { success: false, error: `Failed to insert image: ${e.message}` };
    }

    await saveDocument();
    return { success: true };
}

async function cmdDeleteNode({ type, index }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!type) return { success: false, error: 'Node type is required' };
    if (index === undefined) return { success: false, error: 'Node index is required' };

    const nodes = activeEditor.getNodesOfType(type);
    if (!nodes || nodes.length === 0) {
        return { success: false, error: `No ${type} nodes found in document` };
    }

    const idx = parseInt(index, 10);
    if (idx < 0 || idx >= nodes.length) {
        return { success: false, error: `Index ${index} out of range (${nodes.length} ${type} nodes found)` };
    }

    const { pos, node } = nodes[idx];

    setDocumentMode('suggesting');
    activeEditor.view.focus();
    activeEditor.commands.setTextSelection({ from: pos, to: pos + node.nodeSize });
    activeEditor.commands.deleteSelection();

    await saveDocument();
    setDocumentMode('editing');
    debug(`deleteNode: deleted ${type} at index ${index}`);
    return { success: true, result: { deletedType: type, deletedIndex: index } };
}

async function cmdInsertTable({ rows, cols, data, position, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    // Infer dimensions from data if provided
    const tableRows = rows || (data ? data.length : 2);
    const tableCols = cols || (data && data[0] ? data[0].length : 2);

    setAuthorIfProvided(author);
    setDocumentMode('suggesting');

    // Position the cursor if anchor provided
    let insertPos = null;
    if (position?.after || position?.before) {
        const anchor = position.after || position.before;
        const match = findAnchor(activeEditor, anchor);
        if (!match) {
            return { success: false, error: `Anchor text not found: "${anchor}"` };
        }
        insertPos = position.after ? match.to : match.from;
        activeEditor.view.focus();
        activeEditor.commands.setTextSelection({ from: insertPos, to: insertPos });
    }

    const result = activeEditor.commands.insertTable({ rows: tableRows, cols: tableCols });
    if (!result) {
        return { success: false, error: 'Failed to insert table' };
    }

    // Populate cells with data if provided
    if (data && Array.isArray(data)) {
        // Find the newly created table
        const tablesAfter = activeEditor.getNodesOfType('table');
        const newTable = tablesAfter.find(t => t.pos >= (insertPos || 0) - 5)
            || tablesAfter[tablesAfter.length - 1];

        if (newTable) {
            // Get cells only from this specific table
            const tableEnd = newTable.pos + newTable.node.nodeSize;
            const tableCells = activeEditor.getNodesOfType('tableCell')
                .filter(cell => cell.pos >= newTable.pos && cell.pos < tableEnd);

            // Build list of {cellIndex, content} pairs to insert
            const insertions = [];
            let cellIndex = 0;
            for (let row = 0; row < data.length && row < tableRows; row++) {
                const rowData = data[row];
                if (!Array.isArray(rowData)) {
                    cellIndex += tableCols;
                    continue;
                }
                for (let col = 0; col < tableCols; col++) {
                    if (col < rowData.length && rowData[col]) {
                        insertions.push({ cellIndex, content: rowData[col] });
                    }
                    cellIndex++;
                }
            }

            // Insert in REVERSE order so positions don't shift for unprocessed cells
            activeEditor.view.focus();
            for (let i = insertions.length - 1; i >= 0; i--) {
                const { cellIndex: idx, content } = insertions[i];
                if (tableCells[idx]) {
                    const cellInsertPos = tableCells[idx].pos + 2;
                    activeEditor.commands.setTextSelection({ from: cellInsertPos, to: cellInsertPos });
                    activeEditor.commands.insertContent(content);
                }
            }
        }
    }

    await saveDocument();
    setDocumentMode('editing');
    debug(`insertTable: ${tableRows}x${tableCols} table created${data ? ' with data' : ''}`);
    return { success: true, result: { rows: tableRows, cols: tableCols } };
}

async function cmdUndo() {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const result = activeEditor.commands.undo();
    if (result) {
        await saveDocument();
        debug('undo: success');
    }
    return { success: result };
}

async function cmdRedo() {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const result = activeEditor.commands.redo();
    if (result) {
        await saveDocument();
        debug('redo: success');
    }
    return { success: result };
}

/**
 * Insert a table of contents with bookmarks and internal links.
 *
 * The LLM identifies heading entries (positions + levels) and passes them in.
 * This command:
 *   1. Inserts bookmarkStart/bookmarkEnd pairs at each heading (no style change)
 *   2. Builds a proper tableOfContents node with link marks pointing to those bookmarks
 *   3. Inserts the TOC at the specified position
 *
 * @param {Array} entries - [{level: 1-6, from: N, to: M}, ...] heading positions
 * @param {Object} [position] - {after: "text"} or {before: "text"} or omit for beginning
 * @param {string} [title] - TOC title (default: "Table of Contents", "" for none)
 * @param {Object} [style] - {fontFamily, fontSize, color} for TOC styling. Bold + black by default.
 * @param {Object} [author] - {name, email} for track changes attribution
 */
async function cmdInsertTableOfContents({ entries, position, title, style, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    if (!entries || !Array.isArray(entries) || entries.length === 0) {
        return { success: false, error: 'entries is required: [{level: 1-6, from: N, to: M}, ...]' };
    }

    // Validate and read text for each entry from the document
    const doc = activeEditor.state.doc;
    const schema = activeEditor.state.schema;
    const resolvedEntries = [];
    const timestamp = Date.now();
    for (let i = 0; i < entries.length; i++) {
        const { level, from, to } = entries[i];
        if (!level || !Number.isInteger(from) || !Number.isInteger(to)) {
            return { success: false, error: `Entry ${i}: level, from, and to are required integers` };
        }
        if (level < 1 || level > 6) {
            return { success: false, error: `Entry ${i}: level must be 1-6, got ${level}` };
        }
        if (from >= to || from < 0 || to > doc.content.size) {
            return { success: false, error: `Entry ${i}: invalid range from=${from} to=${to} (doc size: ${doc.content.size})` };
        }
        let text = doc.textBetween(from, to, ' ').trim();
        if (!text) {
            return { success: false, error: `Entry ${i}: no text found at range ${from}-${to}` };
        }
        // Prepend numbering marker if the paragraph has one (e.g., "1.", "(A)")
        try {
            const dom = activeEditor.view.nodeDOM(from);
            const marker = dom?.getAttribute?.('data-marker-type');
            if (marker) text = `${marker} ${text}`;
        } catch {}
        const bookmarkName = `_Toc_${i}_${timestamp}`;
        resolvedEntries.push({ level, from, to, text, bookmarkName });
    }

    // Resolve TOC insertion position before modifying the document
    const anchor = position?.after || position?.before;
    let insertPos = 1;
    if (anchor) {
        const match = findAnchor(activeEditor, anchor);
        if (!match) {
            return { success: false, error: `Anchor text not found: "${anchor}"` };
        }
        insertPos = position?.after ? match.to : match.from;
    }

    setAuthorIfProvided(author);
    setDocumentMode('suggesting');

    // Build everything in a single transaction for atomic undo
    const tr = activeEditor.state.tr;

    // Step 1: Insert bookmarks inside each heading paragraph (work backwards to preserve positions)
    // Positions from getNodes: from = before <p>, to = after </p>
    // Insert inside: from+1 = start of text content, to-1 = end of text content
    const sortedForBookmarks = [...resolvedEntries].sort((a, b) => b.from - a.from);
    for (const entry of sortedForBookmarks) {
        const bmEnd = schema.nodes.bookmarkEnd.create({ id: entry.bookmarkName });
        tr.insert(entry.to - 1, bmEnd);
        const bmStart = schema.nodes.bookmarkStart.create({ name: entry.bookmarkName, id: entry.bookmarkName });
        tr.insert(entry.from + 1, bmStart);
    }

    // Step 2: Build TOC node with plain text + link marks (styling applied after insertion via commands)
    const tocParagraphs = [];
    const tocTitle = title !== undefined ? title : 'Table of Contents';
    const linkMarkType = schema.marks.link;

    if (tocTitle) {
        tocParagraphs.push(
            schema.nodes.paragraph.create({}, [schema.text(tocTitle)])
        );
    }

    const entryIndentLevels = [];
    for (const entry of resolvedEntries) {
        entryIndentLevels.push(entry.level);
        const marks = [
            linkMarkType.create({ href: null, anchor: entry.bookmarkName, name: entry.bookmarkName }),
        ];
        tocParagraphs.push(
            schema.nodes.paragraph.create({}, [schema.text(entry.text, marks)])
        );
    }
    const tocNode = schema.nodes.tableOfContents.create(
        { instruction: `TOC \\o "1-${Math.max(...resolvedEntries.map(e => e.level))}"` },
        tocParagraphs
    );

    // Step 3: Insert TOC node (position mapped through bookmark insertions)
    const mappedInsertPos = tr.mapping.map(insertPos);
    tr.insert(mappedInsertPos, tocNode);
    activeEditor.view.dispatch(tr);

    // Step 4: Apply styling and indentation via SuperDoc commands
    // Switch to editing mode — formatting must not create track changes
    setDocumentMode('editing');
    // Helper: re-fetch TOC and compute a child paragraph's position (each command changes the doc)
    const getTocChild = (childIdx) => {
        const nodes = activeEditor.getNodesOfType('tableOfContents');
        if (nodes.length === 0) return null;
        const t = nodes[0];
        let pos = t.pos + 1;
        for (let c = 0; c < childIdx; c++) pos += t.node.child(c).nodeSize;
        return { pos, node: t.node.child(childIdx), totalChildren: t.node.childCount };
    };

    const firstChild = getTocChild(0);
    if (firstChild) {
        // Apply base styling to each paragraph individually
        // (TextSelection cannot span the TOC node, only its inline-content children)
        for (let c = 0; c < firstChild.totalChildren; c++) {
            const p = getTocChild(c);
            if (!p) break;
            selectRange(activeEditor, p.pos + 1, p.pos + p.node.nodeSize - 1);
            if (style?.bold !== false) activeEditor.commands.setBold();
            activeEditor.commands.setColor(style?.color || '#000000');
            if (style?.fontFamily) activeEditor.commands.setFontFamily(style.fontFamily);
            if (style?.fontSize) activeEditor.commands.setFontSize(style.fontSize);
        }

        // Apply larger font size to title
        if (tocTitle && style?.fontSize) {
            const sizeMatch = style.fontSize.match(/^(\d+(?:\.\d+)?)(pt|px)$/);
            if (sizeMatch) {
                const titleSize = `${parseFloat(sizeMatch[1]) + 2}${sizeMatch[2]}`;
                const p = getTocChild(0);
                if (p) {
                    selectRange(activeEditor, p.pos + 1, p.pos + p.node.nodeSize - 1);
                    activeEditor.commands.setFontSize(titleSize);
                }
            }
        }

        // Apply indentation per entry (level 1 = 0.5", level 2 = 1.0", etc.)
        const titleOffset = tocTitle ? 1 : 0;
        for (let i = 0; i < entryIndentLevels.length; i++) {
            const p = getTocChild(i + titleOffset);
            if (p) {
                selectRange(activeEditor, p.pos + 1, p.pos + p.node.nodeSize - 1);
                activeEditor.commands.setTextIndentation(entryIndentLevels[i] * 36);
            }
        }
    }

    await saveDocument();
    setDocumentMode('editing');

    debug(`insertTableOfContents: inserted TOC with ${resolvedEntries.length} entries and bookmarks`);
    return {
        success: true,
        result: {
            entriesCount: resolvedEntries.length,
            entries: resolvedEntries.map(e => ({ level: e.level, text: e.text, bookmark: e.bookmarkName }))
        }
    };
}

/**
 * Delete the table of contents node from the document.
 * Optionally also removes the bookmarks that were created for TOC entries.
 */
async function cmdDeleteTableOfContents({ removeBookmarks } = {}) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const doc = activeEditor.state.doc;
    const tr = activeEditor.state.tr;

    // Collect all deletions in one pass, apply in a single transaction
    let tocPos = null;
    let tocNodeSize = null;
    const bookmarksToRemove = [];

    doc.descendants((node, pos) => {
        if (node.type.name === 'tableOfContents' && tocPos === null) {
            tocPos = pos;
            tocNodeSize = node.nodeSize;
        }
        if (removeBookmarks &&
            (node.type.name === 'bookmarkStart' || node.type.name === 'bookmarkEnd') &&
            (node.attrs.name?.startsWith('_Toc_') || node.attrs.id?.startsWith('_Toc_'))) {
            bookmarksToRemove.push({ pos, size: node.nodeSize });
        }
    });

    if (tocPos === null) {
        return { success: false, error: 'No table of contents found in document' };
    }

    // Combine TOC + bookmarks, sort backwards, delete in one transaction
    const deletions = [{ pos: tocPos, size: tocNodeSize }, ...bookmarksToRemove];
    deletions.sort((a, b) => b.pos - a.pos);
    for (const del of deletions) {
        tr.delete(del.pos, del.pos + del.size);
    }
    activeEditor.view.dispatch(tr);

    await saveDocument();
    debug('deleteTableOfContents: success');
    return { success: true };
}

async function cmdAddComment({ search, comment, occurrence, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!search) return { success: false, error: 'Search text is required' };
    if (!comment) return { success: false, error: 'Comment text is required' };

    const { match, error: searchError } = findMatch(activeEditor, search, occurrence);
    if (searchError) {
        return { success: false, error: searchError };
    }

    selectRange(activeEditor, match.from, match.to);

    // Add comment via TipTap command
    const authorName = author?.name || editor.user?.name || 'Claude';
    const authorEmail = author?.email || editor.user?.email || 'claude@anthropic.com';

    const result = activeEditor.commands.addComment({
        content: comment,
        author: authorName,
        authorEmail: authorEmail
    });

    if (!result) {
        return { success: false, error: 'Failed to add comment' };
    }

    await saveDocument();
    debug(`addComment: added comment on "${match.text}"`);
    return { success: true, result: { commentedText: match.text } };
}

// Notify VS Code that the webview is ready
debug('Notifying VS Code that webview is ready');
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
