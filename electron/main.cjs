// Electron main process: windows, menus, and all file reading/writing.
const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const mammoth = require('mammoth');
const { importDocx } = require('./docx-import.cjs');
const { readXlsx, writeXlsx } = require('./xlsx-io.cjs');

// The two programs share this shell; each window is one or the other.
const APPS = {
  word: {
    title: 'My Word',
    page: path.join(__dirname, '..', 'dist', 'word', 'index.html'),
    filter: { name: 'Word Document', extensions: ['docx'] },
    ext: 'docx',
    untitled: 'Document',
    size: { width: 1200, height: 900 },
  },
  excel: {
    title: 'My Excel',
    page: path.join(__dirname, '..', 'dist', 'excel', 'index.html'),
    filter: { name: 'Excel Workbook', extensions: ['xlsx'] },
    ext: 'xlsx',
    untitled: 'Book',
    size: { width: 1400, height: 900 },
  },
};

function kindForFile(filePath) {
  const ext = path.extname(filePath).slice(1).toLowerCase();
  return Object.keys(APPS).find((kind) => APPS[kind].ext === ext);
}

// Per-window document state, keyed by webContents id.
const docs = new Map();
const untitledCount = { word: 0, excel: 0 };

function stateOf(win) {
  return docs.get(win.webContents.id);
}

function displayName(state) {
  return state.filePath ? path.basename(state.filePath) : state.untitledName;
}

function updateTitle(win) {
  const state = stateOf(win);
  win.setTitle(`${displayName(state)}${state.dirty ? ' •' : ''} - ${APPS[state.kind].title}`);
  win.setDocumentEdited(state.dirty);
}

function createWindow(kind, filePath) {
  const app_ = APPS[kind];
  const win = new BrowserWindow({
    ...app_.size,
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
  docs.set(id, {
    kind,
    filePath: null,
    dirty: false,
    untitledName: `${app_.untitled}${++untitledCount[kind]}`,
    closing: false,
  });
  updateTitle(win);
  win.setMenu(menus[kind]);
  win.on('focus', () => Menu.setApplicationMenu(menus[kind]));

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

  win.loadFile(app_.page);
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

// Asks the page for something (the .docx bytes, the workbook model, a print
// layout) and waits for its answer.
let nextRequestId = 1;
const pending = new Map();

function requestFromPage(win, channel) {
  return new Promise((resolve, reject) => {
    const requestId = nextRequestId++;
    pending.set(requestId, { resolve, reject });
    win.webContents.send(channel, requestId);
  });
}

for (const reply of ['docx:bytes', 'sheet:model', 'print:ready']) {
  ipcMain.on(reply, (_event, requestId, result, error) => {
    const waiting = pending.get(requestId);
    if (!waiting) return;
    pending.delete(requestId);
    if (error) waiting.reject(new Error(error));
    else waiting.resolve(result);
  });
}

async function fileBytes(win, kind) {
  if (kind === 'excel') return writeXlsx(await requestFromPage(win, 'sheet:collect'));
  return Buffer.from(await requestFromPage(win, 'docx:collect'));
}

async function saveDocument(win, saveAs) {
  const state = stateOf(win);
  const { ext, filter } = APPS[state.kind];
  let target = state.filePath;
  if (!target || saveAs) {
    const result = await dialog.showSaveDialog(win, {
      title: 'Save As',
      defaultPath: state.filePath || `${state.untitledName}.${ext}`,
      filters: [filter],
    });
    if (result.canceled || !result.filePath) return false;
    target = result.filePath.toLowerCase().endsWith(`.${ext}`) ? result.filePath : `${result.filePath}.${ext}`;
  }

  try {
    const bytes = await fileBytes(win, state.kind);
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
    const payload = stateOf(win).kind === 'excel' ? { model: await readXlsx(buffer) } : await readWordFile(buffer);
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

async function readWordFile(buffer) {
  try {
    return { json: await importDocx(buffer) };
  } catch (err) {
    // Fall back to a plainer conversion that copes with unusual files.
    console.error('docx import failed, using fallback:', err);
    const result = await mammoth.convertToHtml(
      { buffer },
      { styleMap: ['u => u', 'strike => s', "p[style-name='Title'] => h1:fresh"] },
    );
    return { html: result.value, simplified: true };
  }
}

async function showOpenDialog(win) {
  const state = win && stateOf(win);
  const kind = state?.kind || 'word';
  const other = kind === 'word' ? 'excel' : 'word';
  const result = await dialog.showOpenDialog(win, {
    title: 'Open',
    filters: [
      APPS[kind].filter,
      APPS[other].filter,
      { name: 'All Office files', extensions: [APPS.word.ext, APPS.excel.ext] },
    ],
    properties: ['openFile'],
  });
  if (result.canceled || !result.filePaths.length) return;
  openPath(result.filePaths[0], win);
}

// Opens a file in the right program. Reuses `win` if it holds an untouched
// new document of the same kind, like Office does.
function openPath(filePath, win) {
  const kind = kindForFile(filePath);
  if (!kind) {
    dialog.showErrorBox('Could not open', `${filePath}\n\nMy Office opens .docx and .xlsx files.`);
    return;
  }
  const state = win && stateOf(win);
  if (state && state.kind === kind && !state.filePath && !state.dirty) openFileInto(win, filePath);
  else createWindow(kind, filePath);
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
    await preparePrint(win);
    const pdf = await win.webContents.printToPDF({ preferCSSPageSize: true, printBackground: true });
    await fs.writeFile(result.filePath, pdf);
    shell.openPath(result.filePath);
  } catch (err) {
    dialog.showErrorBox('Could not export PDF', err.message);
  }
}

// My Excel prints a laid-out copy of its sheets rather than the screen.
async function preparePrint(win) {
  if (stateOf(win).kind === 'excel') await requestFromPage(win, 'print:prepare');
}

async function print(win) {
  try {
    await preparePrint(win);
    win.webContents.print({ printBackground: true });
  } catch (err) {
    dialog.showErrorBox('Could not print', err.message);
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
      return createWindow(win ? stateOf(win).kind : 'word');
    case 'new-word':
      return createWindow('word');
    case 'new-excel':
      return createWindow('excel');
    case 'cut':
    case 'copy':
    case 'paste':
      return win && win.webContents[action]();
    case 'open':
      return showOpenDialog(win);
    case 'save':
      return win && saveDocument(win, false);
    case 'save-as':
      return win && saveDocument(win, true);
    case 'export-pdf':
      return win && exportPdf(win);
    case 'print':
      return win && print(win);
    default:
      // Editing commands are handled by the editor itself.
      return win && win.webContents.send('editor:command', action);
  }
}

// Each program has its own menu bar; the File and View menus are shared.
const menus = {};

function buildMenus() {
  const act = (action) => (_item, win) => runAction(action, win || BrowserWindow.getFocusedWindow());
  // Shown in the menu, but the page handles the key itself.
  const pageKey = (accelerator) => ({ accelerator, registerAccelerator: false });

  const fileMenu = (kind) => ({
    label: '&File',
    submenu: [
      { label: kind === 'excel' ? '&New Workbook' : '&New Document', accelerator: 'CmdOrCtrl+N', click: act('new') },
      kind === 'excel'
        ? { label: 'New Word &Document', click: act('new-word') }
        : { label: 'New Excel &Workbook', click: act('new-excel') },
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
  });

  const viewMenu = (extra = []) => ({
    label: '&View',
    submenu: [
      ...extra,
      { role: 'zoomIn', label: 'Zoom &In', accelerator: 'CmdOrCtrl+=' },
      { role: 'zoomOut', label: 'Zoom &Out' },
      { role: 'resetZoom', label: '&100%' },
      { type: 'separator' },
      { role: 'togglefullscreen', label: '&Full Screen' },
      ...(app.isPackaged ? [] : [{ role: 'toggleDevTools' }]),
    ],
  });

  menus.word = Menu.buildFromTemplate([
    fileMenu('word'),
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
    viewMenu(),
  ]);

  menus.excel = Menu.buildFromTemplate([
    fileMenu('excel'),
    {
      label: '&Edit',
      submenu: [
        { label: '&Undo', accelerator: 'CmdOrCtrl+Z', click: act('undo') },
        { label: '&Redo', accelerator: 'CmdOrCtrl+Y', click: act('redo') },
        { type: 'separator' },
        { role: 'cut', label: 'Cu&t' },
        { role: 'copy', label: '&Copy' },
        { role: 'paste', label: '&Paste' },
        { type: 'separator' },
        { label: 'Fill &Down', ...pageKey('CmdOrCtrl+D'), click: act('fill-down') },
        { label: 'Fill &Right', ...pageKey('CmdOrCtrl+R'), click: act('fill-right') },
        {
          label: 'C&lear',
          submenu: [
            { label: '&Contents', ...pageKey('Delete'), click: act('clear-contents') },
            { label: '&Formats', click: act('clear-formats') },
            { label: '&All', click: act('clear-all') },
          ],
        },
        { type: 'separator' },
        { label: 'Select &All', ...pageKey('CmdOrCtrl+A'), click: act('select-all') },
        { label: '&Find…', accelerator: 'CmdOrCtrl+F', click: act('find') },
        { label: 'R&eplace…', ...pageKey('CmdOrCtrl+H'), click: act('replace') },
        { label: '&Go To…', ...pageKey('CmdOrCtrl+G'), click: act('go-to') },
      ],
    },
    {
      label: '&Insert',
      submenu: [
        { label: '&Rows Above', click: act('insert-rows') },
        { label: '&Columns Left', click: act('insert-cols') },
        { label: '&Sheet', ...pageKey('Shift+F11'), click: act('insert-sheet') },
        { type: 'separator' },
        { label: '&PivotTable…', click: act('insert-pivot') },
        { type: 'separator' },
        { label: '&AutoSum', ...pageKey('Alt+='), click: act('autosum') },
        { label: '&Function…', ...pageKey('Shift+F3'), click: act('insert-function') },
        { label: "&Today's Date", ...pageKey('CmdOrCtrl+;'), click: act('insert-date') },
      ],
    },
    {
      label: 'F&ormat',
      submenu: [
        { label: '&Bold', ...pageKey('CmdOrCtrl+B'), click: act('bold') },
        { label: '&Italic', ...pageKey('CmdOrCtrl+I'), click: act('italic') },
        { label: '&Underline', ...pageKey('CmdOrCtrl+U'), click: act('underline') },
        { type: 'separator' },
        { label: '&Custom Number Format…', click: act('custom-format') },
        { label: 'Column &Width…', click: act('column-width') },
        { label: '&AutoFit Column Width', click: act('autofit-columns') },
        { label: 'Row &Height…', click: act('row-height') },
        { type: 'separator' },
        { label: 'Hide Columns', click: act('hide-cols') },
        { label: 'Unhide Columns', click: act('unhide-cols') },
        { label: 'Hide Rows', click: act('hide-rows') },
        { label: 'Unhide Rows', click: act('unhide-rows') },
        { type: 'separator' },
        { label: '&Rename Sheet', click: act('rename-sheet') },
        { label: '&Delete Sheet', click: act('delete-sheet') },
      ],
    },
    {
      label: '&Data',
      submenu: [
        { label: 'Sort &A to Z', click: act('sort-asc') },
        { label: 'Sort &Z to A', click: act('sort-desc') },
        { type: 'separator' },
        { label: '&PivotTable…', click: act('insert-pivot') },
        { label: '&Refresh PivotTable', click: act('refresh-pivot') },
        { label: 'Refresh A&ll PivotTables', click: act('refresh-all-pivots') },
      ],
    },
    viewMenu([
      {
        label: 'F&reeze Panes',
        submenu: [
          { label: '&Freeze Panes (above and left of the selected cell)', click: act('freeze-panes') },
          { label: 'Freeze &Top Row', click: act('freeze-top-row') },
          { label: 'Freeze First &Column', click: act('freeze-first-col') },
          { label: '&Unfreeze Panes', click: act('unfreeze') },
        ],
      },
      { type: 'separator' },
    ]),
  ]);
  Menu.setApplicationMenu(menus.word);
}

// Double-clicking a .docx or .xlsx (or "Open with") passes the path on the
// command line. The My Excel shortcut passes --app=excel.
function fileFromArgs(argv) {
  return argv.slice(1).find((a) => !a.startsWith('-') && kindForFile(a));
}

function appFromArgs(argv) {
  const flag = argv.find((a) => a.startsWith('--app='));
  const kind = flag && flag.slice('--app='.length);
  return APPS[kind] ? kind : null;
}

function launch(argv, cwd) {
  const file = fileFromArgs(argv);
  if (file) createWindow(kindForFile(file), path.resolve(cwd || process.cwd(), file));
  else createWindow(appFromArgs(argv) || 'word');
}

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_event, argv, cwd) => launch(argv, cwd));

  app.whenReady().then(() => {
    buildMenus();
    launch(process.argv);
  });

  app.on('window-all-closed', () => app.quit());
}
