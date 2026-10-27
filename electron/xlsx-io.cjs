// Reads and writes .xlsx files for My Excel, using ExcelJS.
//
// The spreadsheet page and this module exchange a plain "workbook model":
//   {
//     activeSheet: 0,
//     names: [{ name, ref }],                    // workbook-level named ranges, e.g. ref "Sheet1!$A$1:$B$4"
//     sheets: [{
//       name,
//       cells: [{ r, c, input, value, style }], // 0-based row/column
//       colWidths: { [col]: px }, rowHeights: { [row]: px },
//       frozen: { rows, cols },
//     }],
//   }
// `input` is what the formula engine takes: a number, a boolean, a formula
// starting with "=", or text starting with "'" (the apostrophe keeps text like
// "00123" from being read as a number). `value` is a formula's last result.
//
// Kept on open: values, formulas, number formats, fonts, fills, borders,
// alignment, column widths, row heights, frozen panes, named ranges.
// Dropped: charts, pictures, merged cells, conditional formatting, data
// validation, comments, pivot tables and macros.
const ExcelJS = require('exceljs');

const DEFAULT_COL_PX = 64; // Excel's 8.43-character default
const DEFAULT_ROW_PX = 20; // Excel's 15-point default
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86400000;

// Colours of the default Office theme, in the order Excel numbers them.
const THEME_COLORS = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
// Excel's legacy indexed palette, for the handful of indexes files actually use.
const INDEXED_COLORS = { 8: '000000', 9: 'FFFFFF', 10: 'FF0000', 11: '00FF00', 12: '0000FF', 13: 'FFFF00', 14: 'FF00FF', 15: '00FFFF', 64: '000000' };

// Functions newer than Excel 2007 must be written with a prefix, or Excel shows #NAME?.
const FUTURE_FUNCTIONS = new Set(
  (
    'ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CHOOSECOLS CHOOSEROWS ' +
    'COMBINA CONCAT COT COTH CSC CSCH DAYS DECIMAL DROP ENCODEURL EXPAND FILTERXML FORMULATEXT GAMMA GAUSS HSTACK ' +
    'IFNA IFS IMAGE ISFORMULA ISOWEEKNUM LAMBDA LET MAXIFS MINIFS NUMBERVALUE PDURATION PERMUTATIONA PHI RANDARRAY ' +
    'RRI SEC SECH SEQUENCE SHEET SHEETS SORTBY SWITCH TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW ' +
    'UNICHAR UNICODE UNIQUE VALUETOTEXT VSTACK WEBSERVICE WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR'
  ).split(' '),
);
const WORKSHEET_FUNCTIONS = new Set(['FILTER', 'SORT']);

// ---------- Reading ----------

async function readXlsx(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const sheets = [];
  let activeSheet = 0;
  wb.eachSheet((ws) => {
    if (ws.state === 'veryHidden') return;
    const view = ws.views?.[0] || {};
    if (view.tabSelected) activeSheet = sheets.length;
    sheets.push(readSheet(ws, view));
  });
  if (!sheets.length) sheets.push(emptySheet('Sheet1'));

  const names = [];
  for (const { name, ranges } of wb.definedNames.model || []) {
    // Print areas and other built-in names start with "_xlnm." and aren't for formulas.
    if (ranges?.length === 1 && !name.startsWith('_xlnm.')) names.push({ name, ref: ranges[0] });
  }
  return { activeSheet, names, sheets };
}

function readSheet(ws, view) {
  const sheet = emptySheet(ws.name);
  if (view.state === 'frozen') sheet.frozen = { rows: view.ySplit || 0, cols: view.xSplit || 0 };

  (ws.columns || []).forEach((col, c) => {
    if (col.width && Math.abs(col.width - 8.43) > 0.01) sheet.colWidths[c] = Math.round(col.width * 7 + 5);
    if (col.hidden) sheet.colWidths[c] = 0;
  });

  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const r = rowNumber - 1;
    if (row.hidden) sheet.rowHeights[r] = 0;
    else if (row.height && row.customHeight !== false && Math.abs(row.height - 15) > 0.01)
      sheet.rowHeights[r] = Math.round((row.height * 4) / 3);

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (cell.type === ExcelJS.ValueType.Merge) return; // shown by the merge's top-left cell
      const entry = { r, c: colNumber - 1 };
      const content = readContent(cell);
      if (content) Object.assign(entry, content);
      const style = readStyle(cell);
      if (style) entry.style = style;
      if (content || style) sheet.cells.push(entry);
    });
  });
  return sheet;
}

function readContent(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  switch (cell.type) {
    case ExcelJS.ValueType.Number:
      return { input: v };
    case ExcelJS.ValueType.Boolean:
      return { input: v };
    case ExcelJS.ValueType.Date:
      return { input: dateToSerial(v) };
    case ExcelJS.ValueType.String:
    case ExcelJS.ValueType.SharedString:
      return v === '' ? null : { input: `'${v}` };
    case ExcelJS.ValueType.RichText:
      return { input: `'${v.richText.map((run) => run.text).join('')}` };
    case ExcelJS.ValueType.Hyperlink:
      return { input: `'${typeof v.text === 'string' ? v.text : (v.text?.richText || []).map((t) => t.text).join('')}` };
    case ExcelJS.ValueType.Error:
      return { input: v.error === '#N/A' ? '=NA()' : `'${v.error}` };
    case ExcelJS.ValueType.Formula: {
      const formula = cell.formula; // ExcelJS expands shared formulas for each cell
      if (!formula) return null;
      const out = { input: `=${stripPrefixes(formula)}` };
      const result = cell.result;
      if (result instanceof Date) out.value = dateToSerial(result);
      else if (result && typeof result === 'object' && 'error' in result) out.value = { error: result.error };
      else if (result !== undefined && result !== null) out.value = result;
      return out;
    }
    default:
      return null;
  }
}

function readStyle(cell) {
  const style = {};
  const font = cell.font || {};
  if (font.bold) style.bold = true;
  if (font.italic) style.italic = true;
  if (font.underline && font.underline !== 'none') style.underline = true;
  if (font.strike) style.strike = true;
  if (font.name && font.name !== 'Calibri') style.font = font.name;
  if (font.size && font.size !== 11) style.size = font.size;
  const color = readColor(font.color);
  if (color && color !== '#000000') style.color = color;

  const fill = cell.fill;
  if (fill?.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
    const bg = readColor(fill.pattern === 'solid' ? fill.fgColor : fill.bgColor || fill.fgColor);
    if (bg) style.fill = bg;
  }

  const al = cell.alignment || {};
  if (['left', 'center', 'right'].includes(al.horizontal)) style.align = al.horizontal;
  else if (al.horizontal === 'centerContinuous') style.align = 'center';
  if (al.vertical === 'top' || al.vertical === 'middle') style.valign = al.vertical;
  if (al.wrapText) style.wrap = true;

  if (cell.numFmt && cell.numFmt !== 'General') style.numFmt = cell.numFmt;

  const border = {};
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const s = cell.border?.[side]?.style;
    if (s) border[side] = s;
  }
  if (Object.keys(border).length) style.border = border;

  return Object.keys(style).length ? style : null;
}

function readColor(color) {
  if (!color) return null;
  let hex = null;
  if (typeof color.argb === 'string' && color.argb.length >= 6) hex = color.argb.slice(-6);
  else if (typeof color.theme === 'number') hex = THEME_COLORS[color.theme];
  else if (typeof color.indexed === 'number') hex = INDEXED_COLORS[color.indexed];
  if (!hex) return null;
  if (color.tint) hex = applyTint(hex, color.tint);
  return `#${hex.toLowerCase()}`;
}

// Excel's tint: negative darkens, positive lightens toward white.
function applyTint(hex, tint) {
  return [0, 2, 4]
    .map((i) => {
      const v = parseInt(hex.slice(i, i + 2), 16);
      const t = tint < 0 ? v * (1 + tint) : v + (255 - v) * tint;
      return Math.round(Math.min(255, Math.max(0, t))).toString(16).padStart(2, '0');
    })
    .join('');
}

function dateToSerial(date) {
  return (date.getTime() - EXCEL_EPOCH) / MS_PER_DAY;
}

function stripPrefixes(formula) {
  return mapOutsideStrings(formula, (part) => part.replace(/_xlfn\.|_xlws\.|_xlpm\./gi, ''));
}

function emptySheet(name) {
  return { name, cells: [], colWidths: {}, rowHeights: {}, frozen: { rows: 0, cols: 0 } };
}

// ---------- Writing ----------

async function writeXlsx(model) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'My Excel';
  wb.created = new Date();
  // Ask Excel and LibreOffice to recalculate on open, so results never go stale.
  wb.calcProperties.fullCalcOnLoad = true;

  model.sheets.forEach((sheet, index) => {
    const frozen = sheet.frozen || {};
    const views = [
      {
        state: frozen.rows || frozen.cols ? 'frozen' : 'normal',
        xSplit: frozen.cols || 0,
        ySplit: frozen.rows || 0,
        tabSelected: index === (model.activeSheet || 0),
      },
    ];
    const ws = wb.addWorksheet(sheet.name, { views });

    for (const [c, px] of Object.entries(sheet.colWidths || {})) {
      const col = ws.getColumn(Number(c) + 1);
      if (px === 0) col.hidden = true;
      else col.width = Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);
    }
    for (const [r, px] of Object.entries(sheet.rowHeights || {})) {
      const row = ws.getRow(Number(r) + 1);
      if (px === 0) row.hidden = true;
      else row.height = Math.round(px * 0.75 * 100) / 100;
    }

    for (const cell of sheet.cells) {
      const target = ws.getCell(cell.r + 1, cell.c + 1);
      const value = writeContent(cell);
      if (value !== null) target.value = value;
      if (cell.style) writeStyle(target, cell.style);
    }
  });

  const definedNames = [];
  for (const { name, ref } of model.names || []) definedNames.push({ name, ranges: [ref.replace(/^=/, '')] });
  if (definedNames.length) wb.definedNames.model = definedNames;

  return Buffer.from(await wb.xlsx.writeBuffer());
}

function writeContent(cell) {
  const { input, value } = cell;
  if (input === null || input === undefined || input === '') return null;
  if (typeof input === 'number' || typeof input === 'boolean') return input;
  if (input.startsWith('=')) {
    const out = { formula: addPrefixes(input.slice(1)) };
    if (value && typeof value === 'object' && 'error' in value) out.result = { error: value.error };
    else if (value !== undefined && value !== null) out.result = value;
    return out;
  }
  return input.startsWith("'") ? input.slice(1) : input;
}

function writeStyle(target, style) {
  const font = {};
  if (style.bold) font.bold = true;
  if (style.italic) font.italic = true;
  if (style.underline) font.underline = true;
  if (style.strike) font.strike = true;
  if (style.font) font.name = style.font;
  if (style.size) font.size = style.size;
  if (style.color) font.color = { argb: argb(style.color) };
  if (Object.keys(font).length) target.font = { name: 'Calibri', size: 11, ...font };

  if (style.fill) target.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(style.fill) } };

  const alignment = {};
  if (style.align) alignment.horizontal = style.align;
  if (style.valign) alignment.vertical = style.valign;
  if (style.wrap) alignment.wrapText = true;
  if (Object.keys(alignment).length) target.alignment = alignment;

  if (style.numFmt) target.numFmt = style.numFmt;

  if (style.border) {
    const border = {};
    for (const [side, s] of Object.entries(style.border)) if (s) border[side] = { style: s, color: { argb: 'FF000000' } };
    target.border = border;
  }
}

function argb(hex) {
  return `FF${hex.replace('#', '').toUpperCase()}`;
}

function addPrefixes(formula) {
  return mapOutsideStrings(formula, (part) =>
    part.replace(/(?<![\w.])([A-Za-z][A-Za-z0-9.]*)(?=\s*\()/g, (name) => {
      const upper = name.toUpperCase();
      if (WORKSHEET_FUNCTIONS.has(upper)) return `_xlfn._xlws.${name}`;
      if (FUTURE_FUNCTIONS.has(upper) || (upper.includes('.') && upper !== 'ERROR.TYPE')) return `_xlfn.${name}`;
      return name;
    }),
  );
}

// Applies `fn` to the parts of a formula outside "string literals" and 'quoted sheet names'.
function mapOutsideStrings(formula, fn) {
  return formula
    .split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/)
    .map((part, i) => (i % 2 ? part : fn(part)))
    .join('');
}

module.exports = { readXlsx, writeXlsx, addPrefixes, stripPrefixes, DEFAULT_COL_PX, DEFAULT_ROW_PX };
