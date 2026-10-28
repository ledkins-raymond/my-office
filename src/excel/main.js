import { HyperFormula } from 'hyperformula';
import { Workbook, parseRef, rangeName, cellName, normRange, DEFAULT_COL_PX, DEFAULT_ROW_PX } from './workbook.js';
import { Grid } from './grid.js';
import { NUMBER_FORMATS, formatValue, adjustDecimals, isError } from './format.js';
import { AGGREGATES, sourceFields, isNumericField, writePivot, refreshPivot, pivotOnSheet } from './pivot.js';

const office = window.office;
const $ = (id) => document.getElementById(id);

const wb = new Workbook();
const grid = new Grid($('grid'), wb, {
  onSelection: () => {
    updateNameBox();
    if (!grid.edit) updateFormulaBar();
    updateRibbon();
    updateStats();
  },
  onEdit: (text) => {
    if (document.activeElement !== $('formula')) $('formula').value = text;
    sizeFormulaBar();
    updateAutocomplete();
  },
  onEditEnd: () => {
    hideAutocomplete();
    updateFormulaBar();
  },
  onContextMenu: (e, hit) => showCellMenu(e, hit),
  onShortcut: (e) => handleShortcut(e),
  onSheetStep: (dir) => stepSheet(dir),
  onClipCleared: () => (clipboard = null),
  onAutocompleteKey: (e) => autocompleteKey(e),
  onAutocompleteAccept: (key) => key === 'Tab' && acceptAutocomplete(),
  onAutocompleteCancel: () => {
    if ($('autocomplete').hidden) return false;
    hideAutocomplete();
    return true;
  },
});

// ---------- Dirty tracking ----------
let dirty = false;
function setDirty(value) {
  if (dirty === value) return;
  dirty = value;
  office?.setDirty(value);
}

wb.onChange((kind) => {
  if (kind === 'edit') setDirty(true);
  renderTabs();
  updateFormulaBar();
  updateRibbon();
  updateStats();
  updateNameBox();
});

function status(message) {
  $('status-message').textContent = message;
}

// ---------- File operations (main process does the disk work) ----------
office?.onLoad(({ model, fileName }) => {
  grid.cancelEdit();
  clipboard = null;
  grid.clip = null;
  grid.views.clear();
  grid.shownSheet = undefined;
  wb.load(model);
  setDirty(false);
  status(`Opened ${fileName}`);
  grid.input.focus();
});

office?.onCollectModel(() => {
  grid.commitEdit();
  return wb.toModel();
});

office?.onSaved(({ fileName }) => {
  setDirty(false);
  status(`Saved ${fileName}`);
});

office?.onPreparePrint(() => {
  grid.commitEdit();
  buildPrintArea();
});

// Commands from the menu bar (and some from the ribbon).
const COMMANDS = {
  undo: () => {
    if (grid.edit) return document.execCommand('undo');
    if (!wb.undo()) status('Nothing to undo');
  },
  redo: () => {
    if (grid.edit) return document.execCommand('redo');
    if (!wb.redo()) status('Nothing to redo');
  },
  'select-all': () => {
    if (grid.edit || isTextField(document.activeElement)) return document.activeElement.select?.();
    grid.selectAll();
  },
  find: () => openFind(false),
  replace: () => openFind(true),
  bold: () => toggleFlag('bold'),
  italic: () => toggleFlag('italic'),
  underline: () => toggleFlag('underline'),
  strike: () => toggleFlag('strike'),
  wrap: () => toggleFlag('wrap'),
  'apply-color': () => applyStyle({ color: $('color').value === '#000000' ? null : $('color').value }),
  'apply-fill': () => applyStyle({ fill: $('fill').value }),
  'no-fill': () => applyStyle({ fill: null }),
  'fmt-currency': () => applyStyle({ numFmt: '$#,##0.00' }),
  'fmt-percent': () => applyStyle({ numFmt: '0%' }),
  'fmt-comma': () => applyStyle({ numFmt: '#,##0.00' }),
  'decimals-more': () => changeDecimals(1),
  'decimals-less': () => changeDecimals(-1),
  'fill-down': () => fillFromEdge('down'),
  'fill-right': () => fillFromEdge('right'),
  'insert-rows': () => insertRows(),
  'insert-cols': () => insertCols(),
  'delete-rows': () => deleteRows(),
  'delete-cols': () => deleteCols(),
  'clear-contents': () => wb.clear(wb.activeId, grid.sel, 'contents'),
  'clear-formats': () => wb.clear(wb.activeId, grid.sel, 'formats'),
  'clear-all': () => wb.clear(wb.activeId, grid.sel, 'all'),
  'sort-asc': () => sortSelection(true),
  'sort-desc': () => sortSelection(false),
  autosum: () => autoSum(),
  'insert-sheet': () => wb.addSheet(),
  'delete-sheet': () => deleteSheet(wb.activeId),
  'rename-sheet': () => startRename(wb.activeId),
  'freeze-panes': () => freeze(grid.active.r, grid.active.c),
  'freeze-top-row': () => freeze(1, 0),
  'freeze-first-col': () => freeze(0, 1),
  unfreeze: () => freeze(0, 0),
  'column-width': () => askSize('col'),
  'row-height': () => askSize('row'),
  'autofit-columns': () => grid.autofitColumn(grid.sel.c0),
  'hide-cols': () => wb.setColWidth(wb.activeId, span(grid.sel.c0, grid.sel.c1), 0),
  'unhide-cols': () => unhide('col'),
  'hide-rows': () => wb.setRowHeight(wb.activeId, span(grid.sel.r0, grid.sel.r1), 0),
  'unhide-rows': () => unhide('row'),
  'insert-date': () => typeIntoCell(new Date().toLocaleDateString('en-US')),
  'insert-function': () => insertFunction(),
  'custom-format': () => customFormat(),
  'go-to': () => $('name-box').focus(),
  'insert-pivot': () => openPivotDialog(),
  'refresh-pivot': () => refreshActivePivot(),
  'refresh-all-pivots': () => refreshAllPivots(),
};

function runCommand(name) {
  const fn = COMMANDS[name];
  if (!fn) return;
  // Formatting commands keep an in-progress edit; anything that changes cells commits it first.
  if (grid.edit && !['undo', 'redo', 'select-all'].includes(name)) grid.commitEdit();
  fn();
  if (!$('find-bar').contains(document.activeElement) && document.activeElement !== $('name-box')) grid.input.focus();
}

office?.onCommand((name) => runCommand(name));

function isTextField(node) {
  return node && (node.tagName === 'INPUT' || node.tagName === 'TEXTAREA') && node !== grid.input;
}

// ---------- Ribbon ----------
for (let size of [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]) {
  $('size').append(new Option(String(size), String(size)));
}
for (const { label, fmt } of NUMBER_FORMATS) $('numfmt').append(new Option(label, fmt));
$('numfmt').append(new Option('More formats…', '__custom'));
$('numfmt').append(new Option('Custom', '__other'));

$('ribbon').addEventListener('mousedown', (e) => {
  // Keep focus in the grid so a click on the ribbon doesn't end typing.
  if (e.target.closest('button')) e.preventDefault();
});

$('ribbon').addEventListener('click', (e) => {
  const button = e.target.closest('button');
  if (!button) return;
  if (button.dataset.action) office?.action(button.dataset.action);
  else if (button.dataset.cmd) runCommand(button.dataset.cmd);
  else if (button.dataset.align) {
    const current = currentStyle().align;
    applyStyle({ align: current === button.dataset.align ? null : button.dataset.align });
    grid.input.focus();
  }
});

$('font').addEventListener('change', () => {
  applyStyle({ font: $('font').value === 'Calibri' ? null : $('font').value });
  grid.input.focus();
});
$('size').addEventListener('change', () => {
  const size = Number($('size').value);
  applyStyle({ size: size === 11 ? null : size });
  grid.input.focus();
});
$('color').addEventListener('input', () => ($('color-swatch').style.borderBottomColor = $('color').value));
$('color').addEventListener('change', () => runCommand('apply-color'));
$('fill').addEventListener('input', () => ($('fill-swatch').style.borderBottomColor = $('fill').value));
$('fill').addEventListener('change', () => runCommand('apply-fill'));
$('borders').addEventListener('change', () => {
  applyBorders($('borders').value);
  $('borders').value = '';
  grid.input.focus();
});
$('numfmt').addEventListener('change', () => {
  const value = $('numfmt').value;
  if (value === '__custom') customFormat();
  else if (value !== '__other') applyStyle({ numFmt: value === 'General' ? null : value });
  updateRibbon();
  grid.input.focus();
});

function currentStyle() {
  return wb.getStyle(wb.activeId, grid.active.r, grid.active.c) || {};
}

function updateRibbon() {
  const style = currentStyle();
  for (const button of document.querySelectorAll('[data-flag]')) {
    button.classList.toggle('active', !!style[button.dataset.flag]);
  }
  for (const button of document.querySelectorAll('[data-align]')) {
    button.classList.toggle('active', style.align === button.dataset.align);
  }
  $('font').value = style.font || 'Calibri';
  if (![...$('font').options].some((o) => o.value === $('font').value)) {
    $('font').append(new Option(style.font, style.font));
    $('font').value = style.font;
  }
  const size = String(style.size || 11);
  if (![...$('size').options].some((o) => o.value === size)) $('size').append(new Option(size, size));
  $('size').value = size;
  const fmt = wb.getNumFmt(wb.activeId, grid.active.r, grid.active.c);
  const known = [...$('numfmt').options].find((o) => o.value === fmt);
  $('numfmt').value = known ? fmt : '__other';
}

function applyStyle(patch) {
  wb.setStyle(wb.activeId, grid.sel, patch);
}

function toggleFlag(flag) {
  applyStyle({ [flag]: !currentStyle()[flag] });
}

function changeDecimals(delta) {
  const { r, c } = grid.active;
  const fmt = adjustDecimals(wb.getNumFmt(wb.activeId, r, c), wb.getValue(wb.activeId, r, c), delta);
  applyStyle({ numFmt: fmt });
}

function applyBorders(kind) {
  if (!kind) return;
  const s = grid.sel;
  wb.setStyle(wb.activeId, s, (style, r, c) => {
    const border = { ...style.border };
    const set = (side, value = 'thin') => (border[side] = value);
    if (kind === 'none') return { border: null };
    if (kind === 'all') ['top', 'right', 'bottom', 'left'].forEach((side) => set(side));
    if (kind === 'outside' || kind === 'top') r === s.r0 && set('top');
    if (kind === 'outside' || kind === 'bottom') r === s.r1 && set('bottom');
    if (kind === 'outside' || kind === 'left') c === s.c0 && set('left');
    if (kind === 'outside' || kind === 'right') c === s.c1 && set('right');
    if (kind === 'thick-bottom') r === s.r1 && set('bottom', 'thick');
    return { border: Object.keys(border).length ? border : null };
  });
}

// ---------- Formula bar and name box ----------
function updateNameBox() {
  if (document.activeElement === $('name-box')) return;
  const s = grid.sel;
  const single = s.r0 === s.r1 && s.c0 === s.c1;
  $('name-box').value = single ? cellName(grid.active.r, grid.active.c) : grid.edit ? cellName(grid.active.r, grid.active.c) : rangeName(s);
}

function updateFormulaBar() {
  if (grid.edit) return;
  $('formula').value = grid.inputTextFor(grid.active.r, grid.active.c);
  sizeFormulaBar();
}

function sizeFormulaBar() {
  const lines = $('formula').value.split('\n').length;
  $('formula').style.height = `${Math.min(6, lines) * 18 + 8}px`;
}

$('formula').addEventListener('focus', () => {
  if (!grid.edit) grid.beginEdit('edit', $('formula').value, $('formula'));
  else grid.edit.source = $('formula');
});
$('formula').addEventListener('input', () => {
  grid.setEditText($('formula').value, $('formula'));
  sizeFormulaBar();
  updateAutocomplete();
});
$('formula').addEventListener('keydown', (e) => {
  if (autocompleteKey(e)) return;
  if (e.key === 'Tab' && !$('autocomplete').hidden) {
    e.preventDefault();
    acceptAutocomplete();
    return;
  }
  if (e.key === 'Enter' && e.altKey) {
    e.preventDefault();
    const { selectionStart: a, selectionEnd: b, value } = $('formula');
    $('formula').value = `${value.slice(0, a)}\n${value.slice(b)}`;
    $('formula').setSelectionRange(a + 1, a + 1);
    $('formula').dispatchEvent(new Event('input'));
    return;
  }
  if (e.key === 'Enter' || e.key === 'Tab') {
    e.preventDefault();
    grid.commitEdit();
    if (e.key === 'Tab') grid.moveWithin(0, e.shiftKey ? -1 : 1);
    else grid.moveWithin(e.shiftKey ? -1 : 1, 0);
    grid.input.focus();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    if (!$('autocomplete').hidden) return hideAutocomplete();
    grid.cancelEdit();
    grid.input.focus();
  }
});

$('name-box').addEventListener('focus', () => $('name-box').select());
$('name-box').addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    updateNameBox();
    grid.input.focus();
    return;
  }
  if (e.key !== 'Enter') return;
  e.preventDefault();
  const text = $('name-box').value.trim();
  let target = parseRef(text);
  const named = !target && text && wb.hf.getNamedExpressionFormula(text);
  if (named) {
    const formula = named;
    const m = /^=(?:'?([^'!]+)'?!)?(.+)$/.exec(formula);
    if (m) {
      target = parseRef(m[2]);
      const id = m[1] && wb.hf.getSheetId(m[1]);
      if (target && id !== undefined && id !== null && id !== wb.activeId) wb.setActive(id);
    }
  }
  if (!target) {
    status(`"${text}" is not a cell or range. Try something like B5 or A1:C10.`);
    return;
  }
  grid.commitEdit();
  grid.selectRange(target);
  grid.input.focus();
});

// ---------- Function name suggestions while typing a formula ----------
const FUNCTION_NAMES = HyperFormula.getRegisteredFunctionNames('enGB')
  .filter((name) => !name.startsWith('HF.'))
  .sort();
let suggestions = [];
let suggestionIndex = 0;

function editSource() {
  return grid.edit?.source || grid.input;
}

function tokenAtCaret() {
  const src = editSource();
  const text = src.value;
  if (!text.startsWith('=')) return null;
  const before = text.slice(0, src.selectionStart);
  if ((before.match(/"/g) || []).length % 2) return null; // inside a string
  const m = /(^|[=(,+\-*/^&<>:;\s])([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
  return m ? { start: before.length - m[2].length, text: m[2] } : null;
}

function updateAutocomplete() {
  const token = grid.edit && tokenAtCaret();
  if (!token) return hideAutocomplete();
  const prefix = token.text.toUpperCase();
  suggestions = FUNCTION_NAMES.filter((name) => name.startsWith(prefix)).slice(0, 30);
  if (!suggestions.length || (suggestions.length === 1 && suggestions[0] === prefix)) return hideAutocomplete();
  suggestionIndex = 0;
  const list = $('autocomplete');
  list.replaceChildren(
    ...suggestions.map((name, i) => {
      const li = document.createElement('li');
      li.textContent = name;
      li.className = i === 0 ? 'selected' : '';
      li.addEventListener('mousedown', (e) => {
        e.preventDefault();
        suggestionIndex = i;
        acceptAutocomplete();
      });
      return li;
    }),
  );
  const anchor = editSource().getBoundingClientRect();
  list.style.left = `${anchor.left}px`;
  list.style.top = `${anchor.bottom + 2}px`;
  list.hidden = false;
}

function hideAutocomplete() {
  $('autocomplete').hidden = true;
  suggestions = [];
}

function autocompleteKey(e) {
  if ($('autocomplete').hidden || (e.key !== 'ArrowDown' && e.key !== 'ArrowUp')) return false;
  e.preventDefault();
  suggestionIndex = (suggestionIndex + (e.key === 'ArrowDown' ? 1 : -1) + suggestions.length) % suggestions.length;
  [...$('autocomplete').children].forEach((li, i) => li.classList.toggle('selected', i === suggestionIndex));
  $('autocomplete').children[suggestionIndex]?.scrollIntoView({ block: 'nearest' });
  return true;
}

function acceptAutocomplete() {
  if ($('autocomplete').hidden) return false;
  const token = tokenAtCaret();
  if (!token) return hideAutocomplete();
  const src = editSource();
  const name = `${suggestions[suggestionIndex]}(`;
  const end = src.selectionStart;
  src.value = src.value.slice(0, token.start) + name + src.value.slice(end);
  const caret = token.start + name.length;
  src.setSelectionRange(caret, caret);
  hideAutocomplete();
  if (src === grid.input) grid.input.dispatchEvent(new Event('input'));
  else grid.setEditText(src.value, src);
  return true;
}

// ---------- Keyboard shortcuts (Ctrl+…) ----------
function handleShortcut(e) {
  const k = e.key.toLowerCase();
  const map = {
    b: 'bold',
    i: 'italic',
    u: 'underline',
    5: 'strike',
    d: 'fill-down',
    r: 'fill-right',
    a: 'select-all',
    ';': 'insert-date',
    h: 'replace',
    g: 'go-to',
  };
  if (e.shiftKey && (e.key === '%' || e.code === 'Digit5')) return prevent(e, 'fmt-percent');
  if (e.shiftKey && (e.key === '$' || e.code === 'Digit4')) return prevent(e, 'fmt-currency');
  if (e.shiftKey && (e.key === '!' || e.code === 'Digit1')) return prevent(e, 'fmt-comma');
  if (e.shiftKey && (e.key === '~' || e.code === 'Backquote')) return prevent(e, 'general-format');
  if ((e.key === '+' || e.key === '=') && e.shiftKey) return prevent(e, grid.sel.c0 === 0 ? 'insert-rows' : grid.sel.r0 === 0 ? 'insert-cols' : 'insert-rows');
  if (e.key === '-') return prevent(e, grid.sel.r0 === 0 && grid.sel.c0 !== 0 ? 'delete-cols' : 'delete-rows');
  if (map[k] && !e.shiftKey) return prevent(e, map[k]);
}
COMMANDS['general-format'] = () => applyStyle({ numFmt: null });

function prevent(e, command) {
  e.preventDefault();
  runCommand(command);
}

document.addEventListener('keydown', (e) => {
  if (e.key === '=' && e.altKey) {
    e.preventDefault();
    runCommand('autosum');
  } else if (e.key === 'F11' && e.shiftKey) {
    e.preventDefault();
    runCommand('insert-sheet');
  } else if (e.key === 'F3' && e.shiftKey) {
    e.preventDefault();
    runCommand('insert-function');
  }
});

// ---------- Editing helpers ----------
function typeIntoCell(text) {
  if (grid.edit) {
    const src = editSource();
    const { selectionStart: a, selectionEnd: b, value } = src;
    src.value = value.slice(0, a) + text + value.slice(b);
    src.setSelectionRange(a + text.length, a + text.length);
    src.dispatchEvent(new Event('input'));
  } else {
    wb.enter(wb.activeId, grid.active.r, grid.active.c, text);
  }
}

function fillFromEdge(direction) {
  const s = grid.sel;
  const id = wb.activeId;
  if (direction === 'down') {
    if (s.r0 === s.r1) {
      if (s.r0 === 0) return;
      wb.fill(id, { ...s, r0: s.r0 - 1, r1: s.r0 - 1 }, { ...s, r0: s.r0 - 1 });
    } else wb.fill(id, { ...s, r1: s.r0 }, s);
  } else if (s.c0 === s.c1) {
    if (s.c0 === 0) return;
    wb.fill(id, { ...s, c0: s.c0 - 1, c1: s.c0 - 1 }, { ...s, c0: s.c0 - 1 });
  } else wb.fill(id, { ...s, c1: s.c0 }, s);
}

function wholeRows() {
  return grid.sel.c0 === 0 && grid.sel.c1 >= grid.cols.count - 1;
}

function wholeCols() {
  return grid.sel.r0 === 0 && grid.sel.r1 >= grid.rows.count - 1;
}

function insertRows() {
  wb.shift(wb.activeId, 'row', grid.sel.r0, grid.sel.r1 - grid.sel.r0 + 1);
}

function insertCols() {
  wb.shift(wb.activeId, 'col', grid.sel.c0, grid.sel.c1 - grid.sel.c0 + 1);
}

function deleteRows() {
  if (wholeCols() && !wholeRows()) return status('Select whole rows (click the row numbers) to delete them.');
  wb.shift(wb.activeId, 'row', grid.sel.r0, -(grid.sel.r1 - grid.sel.r0 + 1));
  grid.select(grid.sel.r0, grid.active.c);
}

function deleteCols() {
  if (wholeRows() && !wholeCols()) return status('Select whole columns (click the column letters) to delete them.');
  wb.shift(wb.activeId, 'col', grid.sel.c0, -(grid.sel.c1 - grid.sel.c0 + 1));
  grid.select(grid.active.r, grid.sel.c0);
}

// Sorts the selection, or the block of data around the active cell, by the active column.
function sortSelection(ascending) {
  let range = grid.usedSelection();
  if (range.r0 === range.r1 && range.c0 === range.c1) range = currentRegion(grid.active.r, grid.active.c);
  if (range.r0 === range.r1) return status('Nothing to sort: select the rows to sort first.');
  // Like Excel, treat the first row as headers when it is all text and either
  // bold or above a row that has numbers or dates in it; headers stay on top.
  const id = wb.activeId;
  const col = grid.active.c;
  const cols = span(range.c0, range.c1);
  const firstRow = cols.map((c) => wb.getValue(id, range.r0, c));
  const secondRow = cols.map((c) => wb.getValue(id, range.r0 + 1, c));
  const allText = firstRow.every((v) => typeof v === 'string' || v === null) && firstRow.some((v) => typeof v === 'string');
  const bold = cols.every((c) => wb.getValue(id, range.r0, c) === null || wb.getStyle(id, range.r0, c)?.bold);
  const typesDiffer = secondRow.some((v, i) => v !== null && typeof v !== 'string' && typeof firstRow[i] === 'string');
  const header = allText && (bold || typesDiffer);
  const data = header ? { ...range, r0: range.r0 + 1 } : range;
  wb.sort(id, data, col, ascending);
  status(`Sorted ${data.r1 - data.r0 + 1} rows by column ${rangeName({ r0: 0, c0: col, r1: 0, c1: col }).replace(/\d+/, '')}${header ? ' (kept the header row)' : ''}`);
}

// The block of filled cells around (r, c), like Ctrl+A inside data in Excel.
function currentRegion(r, c) {
  const id = wb.activeId;
  const filled = (rr, cc) => {
    if (rr < 0 || cc < 0) return false;
    const v = wb.getValue(id, rr, cc);
    return v !== null && v !== '' && v !== undefined;
  };
  let region = { r0: r, c0: c, r1: r, c1: c };
  for (let grew = true; grew; ) {
    grew = false;
    const { r0, c0, r1, c1 } = region;
    const rowHas = (rr) => span(c0 - 1, c1 + 1).some((cc) => filled(rr, cc));
    const colHas = (cc) => span(r0 - 1, r1 + 1).some((rr) => filled(rr, cc));
    if (r0 > 0 && rowHas(r0 - 1)) (region.r0 -= 1), (grew = true);
    if (rowHas(r1 + 1)) (region.r1 += 1), (grew = true);
    if (c0 > 0 && colHas(c0 - 1)) (region.c0 -= 1), (grew = true);
    if (colHas(c1 + 1)) (region.c1 += 1), (grew = true);
  }
  return region;
}

// AutoSum: sums the numbers directly above (or to the left of) the active cell.
function autoSum() {
  const id = wb.activeId;
  const s = grid.sel;
  const isNum = (r, c) => r >= 0 && c >= 0 && typeof wb.getValue(id, r, c) === 'number';
  if (s.r0 !== s.r1 || s.c0 !== s.c1) {
    // A selected block: put totals in the row below it.
    wb.transact(() => {
      for (let c = s.c0; c <= s.c1; c++) {
        wb.setInputs(id, s.r1 + 1, c, [[`=SUM(${rangeName({ r0: s.r0, c0: c, r1: s.r1, c1: c })})`]]);
      }
    });
    grid.select(s.r1 + 1, s.c0);
    return;
  }
  const { r, c } = grid.active;
  let range = null;
  if (isNum(r - 1, c)) {
    let top = r - 1;
    while (isNum(top - 1, c)) top--;
    range = { r0: top, c0: c, r1: r - 1, c1: c };
  } else if (isNum(r, c - 1)) {
    let left = c - 1;
    while (isNum(r, left - 1)) left--;
    range = { r0: r, c0: left, r1: r, c1: c - 1 };
  }
  grid.beginEdit('enter', `=SUM(${range ? rangeName(range) : ''})`);
  const pos = range ? grid.input.value.length - 1 : grid.input.value.length - 1;
  grid.input.setSelectionRange(range ? 5 : pos, pos);
  grid.edit.point = range ? { start: 5, end: pos, range, anchor: { r: range.r0, c: range.c0 } } : null;
  $('formula').value = grid.input.value;
}

function insertFunction() {
  if (!grid.edit) grid.beginEdit('enter', '=');
  updateAutocomplete();
}

function freeze(rows, cols) {
  wb.setFrozen(wb.activeId, rows, cols);
  grid.scroller.scrollTop = 0;
  grid.scroller.scrollLeft = 0;
}

function unhide(axis) {
  const id = wb.activeId;
  const s = grid.sel;
  if (axis === 'col') {
    const hidden = span(Math.max(0, s.c0 - 1), s.c1 + 1).filter((c) => wb.colWidth(id, c) === 0);
    if (hidden.length) wb.setColWidth(id, hidden, DEFAULT_COL_PX);
  } else {
    const hidden = span(Math.max(0, s.r0 - 1), s.r1 + 1).filter((r) => wb.rowHeight(id, r) === 0);
    if (hidden.length) wb.setRowHeight(id, hidden, DEFAULT_ROW_PX);
  }
}

async function askSize(axis) {
  const id = wb.activeId;
  const isCol = axis === 'col';
  const current = isCol ? wb.colWidth(id, grid.active.c) : wb.rowHeight(id, grid.active.r);
  // Excel measures widths in characters and heights in points.
  const shown = isCol ? Math.round(((current - 5) / 7) * 100) / 100 : Math.round(current * 0.75 * 100) / 100;
  const result = await ask(isCol ? 'Column width' : 'Row height', [
    { name: 'size', label: isCol ? 'Width (characters, default 8.43)' : 'Height (points, default 15)', value: String(shown) },
  ]);
  const n = Number(result?.size);
  if (!result || !Number.isFinite(n) || n < 0 || n > 409) return;
  const px = isCol ? Math.round(n * 7 + 5) : Math.round(n / 0.75);
  if (isCol) wb.setColWidth(id, span(grid.sel.c0, Math.min(grid.sel.c1, grid.cols.count - 1)), n === 0 ? 0 : px);
  else wb.setRowHeight(id, span(grid.sel.r0, Math.min(grid.sel.r1, grid.rows.count - 1)), n === 0 ? 0 : px);
}

async function customFormat() {
  const current = wb.getNumFmt(wb.activeId, grid.active.r, grid.active.c);
  const result = await ask('Number format', [
    { name: 'fmt', label: 'Excel format code, e.g. 0.000, #,##0, mm/dd/yy, "$"#,##0;[Red]-"$"#,##0', value: current },
  ]);
  if (!result?.fmt) return updateRibbon();
  const sample = formatValue(1234.5678, result.fmt);
  if (!sample && result.fmt !== '@' && result.fmt !== ';;;') return status('That format code was not understood.');
  applyStyle({ numFmt: result.fmt === 'General' ? null : result.fmt });
}

function span(a, b) {
  return Array.from({ length: Math.max(0, b - a + 1) }, (_, i) => a + i);
}

// ---------- Clipboard ----------
// Copying puts tab-separated text (what Excel, Sheets and Word expect) on the
// clipboard. Pasting that same text back in here copies formulas and
// formatting instead, with references adjusted.
let clipboard = null;

function selectionText(range) {
  const id = wb.activeId;
  const rows = [];
  for (let r = range.r0; r <= range.r1; r++) {
    const row = [];
    for (let c = range.c0; c <= range.c1; c++) {
      const text = formatValue(wb.getValue(id, r, c), wb.getNumFmt(id, r, c));
      row.push(/[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text);
    }
    rows.push(row.join('\t'));
  }
  return `${rows.join('\r\n')}\r\n`;
}

function selectionHtml(range) {
  const id = wb.activeId;
  let html = '<table>';
  for (let r = range.r0; r <= range.r1; r++) {
    html += '<tr>';
    for (let c = range.c0; c <= range.c1; c++) html += `<td${cellCss(id, r, c)}>${escapeHtml(formatValue(wb.getValue(id, r, c), wb.getNumFmt(id, r, c)))}</td>`;
    html += '</tr>';
  }
  return `${html}</table>`;
}

function copySelection(data, cut) {
  const range = grid.usedSelection();
  const text = selectionText(range);
  data.setData('text/plain', text);
  data.setData('text/html', selectionHtml(range));
  clipboard = { sheet: wb.activeId, range, cut, text };
  grid.clip = { sheet: wb.activeId, range };
  grid.draw();
  status(cut ? 'Select where to move the cells, then paste' : 'Select where to paste, then press Ctrl+V');
}

grid.input.addEventListener('copy', (e) => {
  if (grid.edit) return;
  e.preventDefault();
  copySelection(e.clipboardData, false);
});

grid.input.addEventListener('cut', (e) => {
  if (grid.edit) return;
  e.preventDefault();
  copySelection(e.clipboardData, true);
});

grid.input.addEventListener('paste', (e) => {
  if (grid.edit) return;
  e.preventDefault();
  const text = e.clipboardData.getData('text/plain');
  const dest = { r: grid.sel.r0, c: grid.sel.c0 };
  if (clipboard && text === clipboard.text && wb.metaOf(clipboard.sheet) && wb.sheetIds.includes(clipboard.sheet)) {
    const target = wb.paste(clipboard, wb.activeId, dest);
    if (clipboard.cut) {
      clipboard = null;
      grid.clip = null;
    }
    grid.selectRange(target, dest);
    status('Ready');
    return;
  }
  if (!text) return;
  const rows = parseTsv(text);
  const target = wb.pasteValues(wb.activeId, dest, rows);
  grid.clip = null;
  clipboard = null;
  grid.selectRange(target, dest);
});

function parseTsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let i = 0;
  let quoted = false;
  text = text.replace(/\r\n/g, '\n').replace(/\n$/, '');
  while (i < text.length) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') (field += '"'), (i += 2);
      else if (ch === '"') (quoted = false), i++;
      else (field += ch), i++;
    } else if (ch === '"' && field === '') (quoted = true), i++;
    else if (ch === '\t') row.push(field), (field = ''), i++;
    else if (ch === '\n') row.push(field), rows.push(row), (row = []), (field = ''), i++;
    else (field += ch), i++;
  }
  row.push(field);
  rows.push(row);
  return rows;
}

function escapeHtml(text) {
  return String(text).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
}

function cellCss(id, r, c) {
  const style = wb.getStyle(id, r, c) || {};
  const value = wb.getValue(id, r, c);
  const css = [];
  if (style.bold) css.push('font-weight:bold');
  if (style.italic) css.push('font-style:italic');
  const deco = [style.underline && 'underline', style.strike && 'line-through'].filter(Boolean).join(' ');
  if (deco) css.push(`text-decoration:${deco}`);
  if (style.color) css.push(`color:${style.color}`);
  if (style.fill) css.push(`background:${style.fill}`);
  if (style.font) css.push(`font-family:'${style.font}'`);
  if (style.size) css.push(`font-size:${style.size}pt`);
  const align = style.align || (typeof value === 'number' ? 'right' : typeof value === 'boolean' || isError(value) ? 'center' : '');
  if (align) css.push(`text-align:${align}`);
  if (style.valign) css.push(`vertical-align:${style.valign}`);
  for (const [side, kind] of Object.entries(style.border || {})) {
    css.push(`border-${side}:${kind === 'thick' ? 2 : kind === 'medium' ? 1.5 : 0.75}pt ${kind === 'double' ? 'double' : 'solid'} #000`);
  }
  return css.length ? ` style="${css.join(';')}"` : '';
}

// ---------- Context menus ----------
function showMenu(e, items) {
  const menu = $('context-menu');
  menu.replaceChildren(
    ...items.map((item) => {
      const li = document.createElement('li');
      if (item === '-') {
        li.className = 'separator';
        return li;
      }
      li.textContent = item.label;
      if (item.key) {
        const key = document.createElement('span');
        key.className = 'key';
        key.textContent = item.key;
        li.append(key);
      }
      li.addEventListener('mousedown', (ev) => ev.preventDefault());
      li.addEventListener('click', () => {
        hideMenu();
        item.run();
      });
      return li;
    }),
  );
  menu.hidden = false;
  const { innerWidth: w, innerHeight: h } = window;
  const rect = menu.getBoundingClientRect();
  menu.style.left = `${Math.min(e.clientX, w - rect.width - 4)}px`;
  menu.style.top = `${Math.min(e.clientY, h - rect.height - 4)}px`;
}

function hideMenu() {
  $('context-menu').hidden = true;
}

window.addEventListener('mousedown', (e) => {
  if (!$('context-menu').contains(e.target)) hideMenu();
  if (!$('autocomplete').contains(e.target) && e.target !== grid.canvas) hideAutocomplete();
});
window.addEventListener('blur', hideMenu);
document.addEventListener('keydown', (e) => e.key === 'Escape' && hideMenu(), true);

function showCellMenu(e, hit) {
  const cmd = (name) => () => runCommand(name);
  const clip = (name) => () => office?.action(name);
  const items = [
    { label: 'Cut', key: 'Ctrl+X', run: clip('cut') },
    { label: 'Copy', key: 'Ctrl+C', run: clip('copy') },
    { label: 'Paste', key: 'Ctrl+V', run: clip('paste') },
    '-',
  ];
  if (hit.kind === 'col') {
    items.push(
      { label: 'Insert columns', run: cmd('insert-cols') },
      { label: 'Delete columns', run: cmd('delete-cols') },
      { label: 'Clear contents', key: 'Delete', run: cmd('clear-contents') },
      '-',
      { label: 'Column width…', run: cmd('column-width') },
      { label: 'AutoFit width', run: cmd('autofit-columns') },
      { label: 'Hide', run: cmd('hide-cols') },
      { label: 'Unhide', run: cmd('unhide-cols') },
    );
  } else if (hit.kind === 'row') {
    items.push(
      { label: 'Insert rows', run: cmd('insert-rows') },
      { label: 'Delete rows', run: cmd('delete-rows') },
      { label: 'Clear contents', key: 'Delete', run: cmd('clear-contents') },
      '-',
      { label: 'Row height…', run: cmd('row-height') },
      { label: 'Hide', run: cmd('hide-rows') },
      { label: 'Unhide', run: cmd('unhide-rows') },
    );
  } else {
    items.push(
      { label: 'Insert rows above', run: cmd('insert-rows') },
      { label: 'Insert columns left', run: cmd('insert-cols') },
      { label: 'Delete rows', run: () => (grid.selectRows(grid.sel.r0, grid.sel.r1), runCommand('delete-rows')) },
      { label: 'Delete columns', run: () => (grid.selectColumns(grid.sel.c0, grid.sel.c1), runCommand('delete-cols')) },
      '-',
      { label: 'Clear contents', key: 'Delete', run: cmd('clear-contents') },
      { label: 'Clear formats', run: cmd('clear-formats') },
      '-',
      { label: 'Sort A to Z', run: cmd('sort-asc') },
      { label: 'Sort Z to A', run: cmd('sort-desc') },
      '-',
      { label: 'Fill down', key: 'Ctrl+D', run: cmd('fill-down') },
      { label: 'Fill right', key: 'Ctrl+R', run: cmd('fill-right') },
      { label: 'Custom number format…', run: cmd('custom-format') },
      { label: 'Freeze panes here', run: cmd('freeze-panes') },
    );
  }
  showMenu(e, items);
}

// ---------- Sheet tabs ----------
let renaming = null;

function renderTabs() {
  if (renaming !== null) return;
  const tabs = $('tabs');
  tabs.replaceChildren(
    ...wb.sheetIds.map((id) => {
      const tab = document.createElement('div');
      tab.className = `tab${id === wb.activeId ? ' active' : ''}`;
      tab.textContent = wb.sheetName(id);
      tab.draggable = true;
      tab.title = 'Double-click to rename, right-click for more';
      tab.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return;
        activate(id);
        // The click itself moves focus to the page; put it back on the grid so typing works.
        setTimeout(() => renaming === null && grid.input.focus());
      });
      tab.addEventListener('dblclick', () => startRename(id));
      tab.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        activate(id);
        const index = wb.sheetIds.indexOf(id);
        showMenu(e, [
          { label: 'Insert sheet', run: () => wb.addSheet() },
          { label: 'Rename', run: () => startRename(id) },
          { label: 'Delete', run: () => deleteSheet(id) },
          '-',
          { label: 'Move left', run: () => index > 0 && wb.moveSheet(id, index - 1) },
          { label: 'Move right', run: () => index < wb.sheetIds.length - 1 && wb.moveSheet(id, index + 1) },
        ]);
      });
      tab.addEventListener('dragstart', (e) => e.dataTransfer.setData('text/sheet', String(id)));
      tab.addEventListener('dragover', (e) => {
        e.preventDefault();
        tab.classList.add('drag-over');
      });
      tab.addEventListener('dragleave', () => tab.classList.remove('drag-over'));
      tab.addEventListener('drop', (e) => {
        e.preventDefault();
        tab.classList.remove('drag-over');
        const moved = Number(e.dataTransfer.getData('text/sheet'));
        if (Number.isFinite(moved) && moved !== id) wb.moveSheet(moved, wb.sheetIds.indexOf(id));
      });
      return tab;
    }),
  );
  tabs.querySelector('.active')?.scrollIntoView({ inline: 'nearest' });
}

function activate(id) {
  if (id === wb.activeId) return;
  grid.commitEdit();
  wb.setActive(id);
  grid.input.focus();
}

function stepSheet(dir) {
  const ids = wb.sheetIds;
  const next = ids[ids.indexOf(wb.activeId) + dir];
  if (next !== undefined) activate(next);
}

function startRename(id) {
  activate(id);
  renderTabs();
  const tab = [...$('tabs').children][wb.sheetIds.indexOf(id)];
  if (!tab) return;
  renaming = id;
  const input = document.createElement('input');
  input.value = wb.sheetName(id);
  tab.replaceChildren(input);
  input.focus();
  input.select();
  let done = false;
  const finish = (save) => {
    if (done) return;
    done = true;
    renaming = null;
    if (save) {
      const error = wb.renameSheet(id, input.value);
      if (error) status(error);
    }
    renderTabs();
    grid.input.focus();
  };
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') finish(true);
    if (e.key === 'Escape') finish(false);
    e.stopPropagation();
  });
  input.addEventListener('blur', () => finish(true));
}

async function deleteSheet(id) {
  if (wb.sheetIds.length < 2) return status('A workbook needs at least one sheet.');
  const used = wb.usedSize(id);
  if (used.rows || used.cols) {
    const ok = await ask(`Delete "${wb.sheetName(id)}"?`, [], 'Delete');
    if (!ok) return;
  }
  wb.removeSheet(id);
  status('Sheet deleted. Undo (Ctrl+Z) brings it back.');
}

$('add-sheet').addEventListener('click', () => {
  grid.commitEdit();
  wb.addSheet();
  grid.input.focus();
});

// ---------- Status bar: sum, average and count of the selection ----------
function updateStats() {
  const s = grid.sel;
  const out = $('selection-stats');
  if (s.r0 === s.r1 && s.c0 === s.c1) return out.replaceChildren();
  const range = grid.usedSelection();
  const id = wb.activeId;
  let count = 0;
  let numbers = 0;
  let sum = 0;
  const cells = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
  if (cells > 500000) return out.replaceChildren();
  for (let r = range.r0; r <= range.r1; r++) {
    for (let c = range.c0; c <= range.c1; c++) {
      const v = wb.getValue(id, r, c);
      if (v === null || v === '' || v === undefined) continue;
      count++;
      if (typeof v === 'number') {
        numbers++;
        sum += v;
      }
    }
  }
  const fmt = wb.getNumFmt(id, grid.active.r, grid.active.c);
  const show = (n) => formatValue(Number(n.toPrecision(15)), fmt === 'General' || fmt === '@' ? 'General' : fmt);
  const parts = [];
  if (numbers) parts.push(`Average: ${show(sum / numbers)}`);
  if (count) parts.push(`Count: ${count}`);
  if (numbers) parts.push(`Sum: ${show(sum)}`);
  out.replaceChildren(...parts.map((text) => Object.assign(document.createElement('span'), { textContent: text })));
}

// ---------- Find and replace ----------
let lastFound = null;

function openFind(withReplace) {
  $('find-bar').hidden = false;
  if (withReplace && $('find-text').value) $('replace-text').focus();
  else $('find-text').focus();
  $('find-text').select();
}

function closeFind() {
  $('find-bar').hidden = true;
  $('find-status').textContent = '';
  grid.input.focus();
}

function matches(text, query) {
  return $('find-case').checked ? text.includes(query) : text.toLowerCase().includes(query.toLowerCase());
}

function findNext() {
  const query = $('find-text').value;
  if (!query) return false;
  const id = wb.activeId;
  const used = wb.usedSize(id);
  const total = used.rows * used.cols;
  if (!total) return notFound();
  const start = grid.active.r * used.cols + grid.active.c;
  for (let i = 1; i <= total; i++) {
    const n = (start + i) % total;
    const r = Math.floor(n / used.cols);
    const c = n % used.cols;
    const text = wb.isFormula(id, r, c) ? wb.getInput(id, r, c) : formatValue(wb.getValue(id, r, c), wb.getNumFmt(id, r, c));
    if (text && matches(text, query)) {
      grid.select(r, c);
      lastFound = { id, r, c };
      $('find-status').textContent = '';
      return true;
    }
  }
  return notFound();
}

function notFound() {
  $('find-status').textContent = 'Not found';
  lastFound = null;
  return false;
}

function replaceIn(text, query, replacement) {
  const flags = $('find-case').checked ? 'g' : 'gi';
  return text.replace(new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags), () => replacement);
}

function replaceOne() {
  const query = $('find-text').value;
  const { r, c } = grid.active;
  const id = wb.activeId;
  const input = grid.inputTextFor(r, c);
  if (lastFound && lastFound.id === id && lastFound.r === r && lastFound.c === c && matches(input, query)) {
    wb.enter(id, r, c, replaceIn(input, query, $('replace-text').value));
  }
  findNext();
}

function replaceAll() {
  const query = $('find-text').value;
  if (!query) return;
  const id = wb.activeId;
  const used = wb.usedSize(id);
  let count = 0;
  wb.transact(() => {
    for (let r = 0; r < used.rows; r++) {
      for (let c = 0; c < used.cols; c++) {
        const input = grid.inputTextFor(r, c);
        if (!input || !matches(input, query)) continue;
        const next = replaceIn(input, query, $('replace-text').value);
        wb.setInputs(id, r, c, [[next === '' ? null : next.startsWith('=') ? next : toTyped(next)]]);
        count++;
      }
    }
  });
  if (!count) wb.undo();
  $('find-status').textContent = `${count} replaced`;
}

function toTyped(text) {
  const n = Number(text);
  return text.trim() !== '' && Number.isFinite(n) ? n : `'${text}`;
}

$('find-text').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') findNext();
  if (e.key === 'Escape') closeFind();
});
$('replace-text').addEventListener('keydown', (e) => e.key === 'Escape' && closeFind());
$('find-next').addEventListener('click', findNext);
$('replace-one').addEventListener('click', replaceOne);
$('replace-all').addEventListener('click', replaceAll);
$('find-close').addEventListener('click', closeFind);

// ---------- Pivot tables ----------
// The data under the cursor (or the selection) is summarised on a new sheet; see pivot.js.

function quoteName(name) {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

// "'Sales Data'!A1:D20" or "A1:D20" (this sheet) → { sheetId, sheetName, range }.
function parseDataRef(text) {
  const m = /^\s*(?:(?:'((?:[^']|'')+)'|([^!]+))!)?([^!]+)$/.exec(text);
  if (!m) return null;
  const name = m[1] ? m[1].replace(/''/g, "'") : m[2]?.trim();
  const sheetId = name ? wb.sheetIds.find((id) => wb.sheetName(id).toLowerCase() === name.toLowerCase()) : wb.activeId;
  const range = parseRef(m[3]);
  if (sheetId === undefined || !range) return null;
  return { sheetId, sheetName: wb.sheetName(sheetId), range };
}

function fillPivotFields() {
  const ref = parseDataRef($('pivot-range').value);
  const error = $('pivot-error');
  const selects = ['pivot-row1', 'pivot-row2', 'pivot-col', 'pivot-val1', 'pivot-val2'].map($);
  if (!ref || ref.range.r1 <= ref.range.r0) {
    error.textContent = 'Enter a range with a header row and data below it, like A1:D50 or \'Sales\'!A1:D50.';
    selects.forEach((s) => s.replaceChildren(new Option('(none)', '')));
    $('pivot-ok').disabled = true;
    return;
  }
  error.textContent = '';
  $('pivot-ok').disabled = false;
  const fields = sourceFields(wb, ref.sheetId, ref.range);
  const keep = selects.map((s) => s.value);
  selects.forEach((s, i) => {
    s.replaceChildren(new Option('(none)', ''), ...fields.map((f) => new Option(f.name, f.name)));
    if (fields.some((f) => f.name === keep[i])) s.value = keep[i];
  });
  // Sensible first guess: rows by the first text column, sum of the last number column.
  if (!keep[0] && !keep[3]) {
    const numeric = fields.filter((f) => isNumericField(wb, ref.sheetId, ref.range, f.col));
    const text = fields.filter((f) => !numeric.includes(f));
    $('pivot-row1').value = (text[0] || fields[0]).name;
    const value = numeric.at(-1) || fields.find((f) => f.name !== $('pivot-row1').value) || fields[0];
    $('pivot-val1').value = value.name;
    $('pivot-agg1').value = numeric.includes(value) ? 'sum' : 'count';
  }
}

function openPivotDialog() {
  let range = grid.usedSelection();
  const single = range.r0 === range.r1 && range.c0 === range.c1;
  if (single) range = currentRegion(grid.active.r, grid.active.c);
  if (range.r0 === range.r1) {
    status('Click inside your data (with a header row on top) first, then choose PivotTable.');
    return;
  }
  for (const id of ['pivot-agg1', 'pivot-agg2']) {
    $(id).replaceChildren(...Object.entries(AGGREGATES).map(([value, a]) => new Option(a.label, value)));
  }
  for (const id of ['pivot-row1', 'pivot-row2', 'pivot-col', 'pivot-val1', 'pivot-val2']) $(id).replaceChildren();
  $('pivot-range').value = `${quoteName(wb.sheetName(wb.activeId))}!${rangeName(range)}`;
  fillPivotFields();
  const dialog = $('pivot-dialog');
  dialog.returnValue = '';
  dialog.showModal();
  $('pivot-row1').focus();
}

$('pivot-range').addEventListener('input', fillPivotFields);

$('pivot-dialog').addEventListener('close', () => {
  const dialog = $('pivot-dialog');
  if (dialog.returnValue !== 'ok') return grid.input.focus();
  const ref = parseDataRef($('pivot-range').value);
  if (!ref) return;
  const rows = [$('pivot-row1').value, $('pivot-row2').value].filter(Boolean);
  const columns = [$('pivot-col').value].filter(Boolean);
  const values = [
    { field: $('pivot-val1').value, agg: $('pivot-agg1').value },
    { field: $('pivot-val2').value, agg: $('pivot-agg2').value },
  ].filter((v) => v.field);
  if (!rows.length && !columns.length && !values.length) {
    status('Pick at least one field for the pivot table.');
    return;
  }
  try {
    grid.commitEdit();
    writePivot(wb, { source: { sheet: ref.sheetName, range: rangeName(ref.range) }, rows: [...new Set(rows)], columns, values });
    grid.select(0, 0);
    status(`Pivot table created on ${wb.sheetName(wb.activeId)}. It updates as the data changes; Refresh picks up new rows and categories.`);
  } catch (err) {
    status(err.message);
  }
  grid.input.focus();
});

function refreshActivePivot() {
  const def = pivotOnSheet(wb, wb.activeId);
  if (!def) {
    if (wb.pivots.length) return refreshAllPivots();
    return status('This sheet has no pivot table. Use PivotTable to make one.');
  }
  try {
    refreshPivot(wb, def);
    status('Pivot table refreshed.');
  } catch (err) {
    status(`Could not refresh: ${err.message}`);
  }
}

function refreshAllPivots() {
  const active = wb.activeId;
  let done = 0;
  const problems = [];
  wb.transact(() => {
    for (const def of wb.pivots) {
      try {
        refreshPivot(wb, def);
        done++;
      } catch (err) {
        problems.push(`${def.sheet}: ${err.message}`);
      }
    }
    wb.activeId = active;
  });
  status(problems.length ? `Refreshed ${done}. ${problems.join(' ')}` : `Refreshed ${done} pivot table${done === 1 ? '' : 's'}.`);
}

// ---------- Dialog ----------
function ask(title, fields, okLabel = 'OK') {
  const dialog = $('prompt');
  $('prompt-title').textContent = title;
  dialog.querySelector('button.primary').textContent = okLabel;
  const container = $('prompt-fields');
  container.replaceChildren(
    ...fields.map((field) => {
      const label = document.createElement('label');
      label.textContent = field.label;
      const input = document.createElement('input');
      input.name = field.name;
      input.value = field.value || '';
      label.append(input);
      return label;
    }),
  );
  return new Promise((resolve) => {
    dialog.addEventListener(
      'close',
      () => {
        if (dialog.returnValue !== 'ok') return resolve(null);
        const out = {};
        for (const input of container.querySelectorAll('input')) out[input.name] = input.value;
        resolve(out);
        grid.input.focus();
      },
      { once: true },
    );
    dialog.returnValue = '';
    dialog.showModal();
    container.querySelector('input')?.select();
  });
}

// ---------- Printing ----------
// Printing and PDF use a plain table of each sheet's used cells, built just before printing.
function buildPrintArea() {
  const area = $('print-area');
  area.replaceChildren();
  for (const id of wb.sheetIds) {
    const used = wb.usedSize(id);
    if (!used.rows || !used.cols) continue;
    const section = document.createElement('section');
    if (wb.sheetIds.length > 1) {
      const h1 = document.createElement('h1');
      h1.textContent = wb.sheetName(id);
      section.append(h1);
    }
    let html = '<table><colgroup>';
    for (let c = 0; c < used.cols; c++) html += `<col style="width:${wb.colWidth(id, c)}px">`;
    html += '</colgroup>';
    for (let r = 0; r < used.rows; r++) {
      const h = wb.rowHeight(id, r);
      if (!h) continue;
      html += `<tr style="height:${h}px">`;
      for (let c = 0; c < used.cols; c++) {
        if (!wb.colWidth(id, c)) {
          html += '<td style="display:none"></td>';
          continue;
        }
        const wrap = wb.getStyle(id, r, c)?.wrap ? ' class="wrap"' : '';
        html += `<td${wrap}${cellCss(id, r, c)}>${escapeHtml(formatValue(wb.getValue(id, r, c), wb.getNumFmt(id, r, c)))}</td>`;
      }
      html += '</tr>';
    }
    section.insertAdjacentHTML('beforeend', `${html}</table>`);
    area.append(section);
  }
  if (!area.children.length) area.textContent = ' ';
}

// ---------- Start ----------
grid.start();
renderTabs();
updateFormulaBar();
updateNameBox();
updateRibbon();
grid.input.focus();
window.myExcel = { wb, grid }; // handy from the developer tools
