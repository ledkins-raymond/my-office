// Electron main process: windows, menus, and all file reading/writing.
const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const mammoth = require('mammoth');
const { importDocx } = require('./docx-import.cjs');

const WORD_PAGE = path.join(__dirname, '..', 'dist', 'word', 'index.html');
const DOCX_FILTER = { name: 'Word Document', extensions: ['docx'] };

// Per-window document state, keyed by webContents id.
const docs = new Map();
let untitledCount = 0;

function stateOf(win) {
  return docs.get(win.webContents.id);
}

function displayName(state) {
  return state.filePath ? path.basename(state.filePath) : state.untitledName;
}

function updateTitle(win) {
  const state = stateOf(win);
  win.setTitle(`${displayName(state)}${state.dirty ? ' •' : ''} - My Word`);
  win.setDocumentEdited(state.dirty);
}

function createWordWindow(filePath) {
  const win = new BrowserWindow({
    width: 1200,
    height: 900,
    minWidth: 700,
    minHeight: 500,
    show: false,
    backgroundColor: '#ffffff', // also the paper colour around printed pages
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true,
    },
  });

  const id = win.webContents.id;
  docs.set(id, { filePath: null, dirty: false, untitledName: `Document${++untitledCount}`, closing: false });
  updateTitle(win);

  win.once('ready-to-show', () => win.show());
  // Keep our "Document1 - My Word" title instead of the page's <title>.
  win.on('page-title-updated', (event) => event.preventDefault());
  win.webContents.once('did-finish-load', () => {
    if (filePath) openFileInto(win, filePath);
  });

  // Links in documents open in the default browser, never inside the app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:|^mailto:/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (event) => event.preventDefault());

  win.on('close', (event) => {
    const state = docs.get(id);
    if (!state.dirty || state.closing) return;
    event.preventDefault();
    confirmDiscard(win).then((ok) => {
      if (ok) {
        state.closing = true;
        win.close();
      }
    });
  });
  win.on('closed', () => docs.delete(id));

  win.loadFile(WORD_PAGE);
  return win;
}

// Asks to save unsaved changes. Resolves true when it is safe to continue.
async function confirmDiscard(win) {
  const { response } = await dialog.showMessageBox(win, {
    type: 'warning',
    buttons: ['Save', "Don't Save", 'Cancel'],
    defaultId: 0,
    cancelId: 2,
    message: `Do you want to save changes to ${displayName(stateOf(win))}?`,
  });
  if (response === 2) return false;
  if (response === 1) return true;
  return saveDocument(win, false);
}

// The renderer turns the editor contents into .docx bytes on request.
let nextRequestId = 1;
const pendingDocx = new Map();

function requestDocxBytes(win) {
  return new Promise((resolve, reject) => {
    const requestId = nextRequestId++;
    pendingDocx.set(requestId, { resolve, reject });
    win.webContents.send('docx:collect', requestId);
  });
}

ipcMain.on('docx:bytes', (_event, requestId, bytes, error) => {
  const pending = pendingDocx.get(requestId);
  if (!pending) return;
  pendingDocx.delete(requestId);
  if (error) pending.reject(new Error(error));
  else pending.resolve(Buffer.from(bytes));
});

async function saveDocument(win, saveAs) {
  const state = stateOf(win);
  let target = state.filePath;
  if (!target || saveAs) {
    const result = await dialog.showSaveDialog(win, {
      title: 'Save As',
      defaultPath: state.filePath || `${state.untitledName}.docx`,
      filters: [DOCX_FILTER],
    });
    if (result.canceled || !result.filePath) return false;
    target = result.filePath.toLowerCase().endsWith('.docx') ? result.filePath : `${result.filePath}.docx`;
  }

  try {
    const bytes = await requestDocxBytes(win);
    await fs.writeFile(target, bytes);
  } catch (err) {
    dialog.showErrorBox('Could not save', `${target}\n\n${err.message}`);
    return false;
  }

  state.filePath = target;
  state.dirty = false;
  app.addRecentDocument(target);
  win.webContents.send('doc:saved', { fileName: path.basename(target) });
  updateTitle(win);
  return true;
}

async function openFileInto(win, filePath) {
  try {
    const buffer = await fs.readFile(filePath);
    let payload;
    try {
      payload = { json: await importDocx(buffer) };
    } catch (err) {
      // Fall back to a plainer conversion that copes with unusual files.
      console.error('docx import failed, using fallback:', err);
      const result = await mammoth.convertToHtml(
        { buffer },
        { styleMap: ['u => u', 'strike => s', "p[style-name='Title'] => h1:fresh"] },
      );
      payload = { html: result.value, simplified: true };
    }
    const state = stateOf(win);
    state.filePath = filePath;
    state.dirty = false;
    app.addRecentDocument(filePath);
    win.webContents.send('doc:load', { ...payload, fileName: path.basename(filePath) });
    updateTitle(win);
  } catch (err) {
    dialog.showErrorBox('Could not open', `${filePath}\n\n${err.message}`);
  }
}

async function showOpenDialog(win) {
  const result = await dialog.showOpenDialog(win, {
    title: 'Open',
    filters: [DOCX_FILTER],
    properties: ['openFile'],
  });
  if (result.canceled || !result.filePaths.length) return;
  const filePath = result.filePaths[0];
  const state = win && stateOf(win);
  // Reuse the window if it holds an untouched new document, like Word does.
  if (state && !state.filePath && !state.dirty) openFileInto(win, filePath);
  else createWordWindow(filePath);
}

async function exportPdf(win) {
  const state = stateOf(win);
  const base = state.filePath ? path.basename(state.filePath, path.extname(state.filePath)) : state.untitledName;
  const result = await dialog.showSaveDialog(win, {
    title: 'Export to PDF',
    defaultPath: path.join(state.filePath ? path.dirname(state.filePath) : app.getPath('documents'), `${base}.pdf`),
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  });
  if (result.canceled || !result.filePath) return;
  try {
    const pdf = await win.webContents.printToPDF({ preferCSSPageSize: true, printBackground: true });
    await fs.writeFile(result.filePath, pdf);
    shell.openPath(result.filePath);
  } catch (err) {
    dialog.showErrorBox('Could not export PDF', err.message);
  }
}

ipcMain.on('doc:dirty', (event, dirty) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  const state = win && stateOf(win);
  if (!state || state.dirty === dirty) return;
  state.dirty = dirty;
  updateTitle(win);
});

ipcMain.on('menu:action', (event, action) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (win) runAction(action, win);
});

function runAction(action, win) {
  switch (action) {
    case 'new':
      return createWordWindow();
    case 'open':
      return showOpenDialog(win);
    case 'save':
      return win && saveDocument(win, false);
    case 'save-as':
      return win && saveDocument(win, true);
    case 'export-pdf':
      return win && exportPdf(win);
    case 'print':
      return win && win.webContents.print({ printBackground: true });
    default:
      // Editing commands are handled by the editor itself.
      return win && win.webContents.send('editor:command', action);
  }
}

function buildMenu() {
  const act = (action) => (_item, win) => runAction(action, win || BrowserWindow.getFocusedWindow());
  const template = [
    {
      label: '&File',
      submenu: [
        { label: '&New', accelerator: 'CmdOrCtrl+N', click: act('new') },
        { label: '&Open…', accelerator: 'CmdOrCtrl+O', click: act('open') },
        { type: 'separator' },
        { label: '&Save', accelerator: 'CmdOrCtrl+S', click: act('save') },
        { label: 'Save &As…', accelerator: 'F12', click: act('save-as') },
        { label: '&Export to PDF…', click: act('export-pdf') },
        { type: 'separator' },
        { label: '&Print…', accelerator: 'CmdOrCtrl+P', click: act('print') },
        { type: 'separator' },
        { role: 'close', label: '&Close' },
        { role: 'quit', label: 'E&xit' },
      ],
    },
    {
      label: '&Edit',
      submenu: [
        { label: '&Undo', accelerator: 'CmdOrCtrl+Z', click: act('undo') },
        { label: '&Redo', accelerator: 'CmdOrCtrl+Y', click: act('redo') },
        { type: 'separator' },
        { role: 'cut', label: 'Cu&t' },
        { role: 'copy', label: '&Copy' },
        { role: 'paste', label: '&Paste' },
        { role: 'pasteAndMatchStyle', label: 'Paste as Plain &Text', accelerator: 'CmdOrCtrl+Shift+V' },
        { type: 'separator' },
        { role: 'selectAll', label: 'Select &All' },
        { label: '&Find…', accelerator: 'CmdOrCtrl+F', click: act('find') },
      ],
    },
    {
      label: '&Insert',
      submenu: [
        { label: '&Table…', click: act('insert-table') },
        { label: '&Picture…', click: act('insert-image') },
        { label: '&Link…', accelerator: 'CmdOrCtrl+K', click: act('insert-link') },
        { label: 'Page &Break', accelerator: 'CmdOrCtrl+Enter', click: act('insert-page-break') },
        { label: '&Horizontal Line', click: act('insert-rule') },
      ],
    },
    {
      label: '&View',
      submenu: [
        { role: 'zoomIn', label: 'Zoom &In', accelerator: 'CmdOrCtrl+=' },
        { role: 'zoomOut', label: 'Zoom &Out' },
        { role: 'resetZoom', label: '&100%' },
        { type: 'separator' },
        { role: 'togglefullscreen', label: '&Full Screen' },
        ...(app.isPackaged ? [] : [{ role: 'toggleDevTools' }]),
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// Double-clicking a .docx (or "Open with") passes the path on the command line.
function fileFromArgs(argv) {
  return argv.slice(1).find((a) => a.toLowerCase().endsWith('.docx') && !a.startsWith('-'));
}

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_event, argv) => {
    const file = fileFromArgs(argv);
    if (file) createWordWindow(path.resolve(file));
    else BrowserWindow.getAllWindows()[0]?.focus();
  });

  app.whenReady().then(() => {
    buildMenu();
    const file = fileFromArgs(process.argv);
    createWordWindow(file ? path.resolve(file) : undefined);
  });

  app.on('window-all-closed', () => app.quit());
}
