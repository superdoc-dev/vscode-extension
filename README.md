# SuperDoc VS Code Extension

SuperDoc in VS Code.

## Setup

```bash
cd superdoc-vscode-extension
npm install
npm run compile
```

## Run in Debug Mode

1. Open this folder in VS Code
2. Press `F5` (or Run > Start Debugging)
3. A new VS Code window opens with the extension loaded
4. Open any `.docx` file to test

## Features

- Edit .docx files directly in VS Code
- Auto-save on changes (1 second debounce)
- Auto-reload when file changes externally (e.g., from another process)
