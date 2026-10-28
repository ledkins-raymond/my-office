// Pivot tables, built from ordinary formulas.
//
// A pivot summarises a block of data (first row = headers) by the distinct
// values of some "row" fields and, optionally, one "column" field. Each
// summary cell is a live SUMIFS / COUNTIFS / AVERAGEIFS formula pointing back
// at the data, so numbers update when the data changes and the result reads
// correctly in Excel, Google Sheets and LibreOffice. It is not an Excel
// PivotTable object: new categories in the data appear only after Refresh.
//
// A definition looks like:
//   { name: 'PivotTable1', sheet: 'PivotTable1',
//     source: { sheet: 'Data', range: 'A1:D200' },
//     rows: ['Region'], columns: ['Quarter'],
//     values: [{ field: 'Amount', agg: 'sum' }] }
// Fields are named by their header text, so moving columns around doesn't break them.
import { parseRef, rangeName, cellName, quoteSheetName } from './workbook.js';

export const AGGREGATES = {
  sum: { label: 'Sum', ifs: 'SUMIFS', all: 'SUM' },
  count: { label: 'Count', ifs: 'COUNTIFS', all: 'COUNTA' },
  // Averages are SUMIFS ÷ COUNTIFS: the formula engine has no AVERAGEIFS.
  average: { label: 'Average', ifs: null, all: 'AVERAGE' },
};

const MAX_ROW_ITEMS = 5000;
const MAX_COLUMN_ITEMS = 200;
const HEADER_FILL = '#dde7f0';
const TOTAL_FILL = '#eef3f8';

// ---------- Reading the data ----------

// Headers of the data block: [{ name, col }] in sheet columns.
export function sourceFields(wb, sheetId, range) {
  const fields = [];
  const seen = new Set();
  for (let c = range.c0; c <= range.c1; c++) {
    const v = wb.getValue(sheetId, range.r0, c);
    let name = v === null || v === undefined || v === '' ? `Column ${c - range.c0 + 1}` : String(v && typeof v === 'object' ? v.error : v);
    while (seen.has(name.toLowerCase())) name += ' (2)';
    seen.add(name.toLowerCase());
    fields.push({ name, col: c });
  }
  return fields;
}

// True when most filled cells of a field are numbers (so Sum makes sense).
export function isNumericField(wb, sheetId, range, col) {
  let numbers = 0;
  let filled = 0;
  for (let r = range.r0 + 1; r <= Math.min(range.r1, range.r0 + 500); r++) {
    const v = wb.getValue(sheetId, r, col);
    if (v === null || v === undefined || v === '') continue;
    filled++;
    if (typeof v === 'number') numbers++;
  }
  return filled > 0 && numbers / filled >= 0.8;
}

// Resolves a definition against the workbook, or throws an Error a person can read.
export function resolve(wb, def) {
  const sheetId = wb.sheetIds.find((id) => wb.sheetName(id).toLowerCase() === def.source.sheet.toLowerCase());
  if (sheetId === undefined) throw new Error(`The data sheet "${def.source.sheet}" no longer exists.`);
  const range = parseRef(def.source.range);
  if (!range || range.r1 <= range.r0) throw new Error('The data needs a header row and at least one row below it.');
  const fields = sourceFields(wb, sheetId, range);
  const find = (name) => {
    const f = fields.find((x) => x.name.toLowerCase() === String(name).toLowerCase());
    if (!f) throw new Error(`The column "${name}" is no longer in the data.`);
    return f;
  };
  const rows = def.rows.map(find);
  const columns = def.columns.map(find);
  const values = def.values.map((v) => ({ ...find(v.field), agg: AGGREGATES[v.agg] ? v.agg : 'sum' }));
  if (!rows.length && !columns.length && !values.length) throw new Error('Pick at least one field.');
  return { sheetId, range, fields, rows, columns, values };
}

// Distinct values of a column, matched without regard to case like Excel, sorted, blanks last.
function distinct(wb, sheetId, range, col) {
  const byKey = new Map();
  for (let r = range.r0 + 1; r <= range.r1; r++) {
    const v = wb.getValue(sheetId, r, col);
    const value = v === '' || v === undefined || (v && typeof v === 'object') ? null : v;
    const k = typeof value === 'string' ? `s:${value.toLowerCase()}` : `${typeof value}:${value}`;
    if (!byKey.has(k)) byKey.set(k, value);
  }
  const rank = (v) => (v === null ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  return [...byKey.values()].sort((a, b) => {
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (typeof a === 'number') return a - b;
    return String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true });
  });
}

// Every combination of row-field values that actually occurs in the data, in sorted order.
function rowCombinations(wb, sheetId, range, rowFields) {
  if (!rowFields.length) return [];
  const orders = rowFields.map((f) => {
    const order = new Map();
    distinct(wb, sheetId, range, f.col).forEach((v, i) => order.set(keyOf(v), i));
    return order;
  });
  const combos = new Map();
  for (let r = range.r0 + 1; r <= range.r1; r++) {
    const values = rowFields.map((f) => {
      const v = wb.getValue(sheetId, r, f.col);
      return v === '' || v === undefined || (v && typeof v === 'object') ? null : v;
    });
    const k = values.map(keyOf).join('\u0001');
    if (!combos.has(k)) combos.set(k, values);
  }
  return [...combos.values()].sort((a, b) => {
    for (let i = 0; i < a.length; i++) {
      const d = orders[i].get(keyOf(a[i])) - orders[i].get(keyOf(b[i]));
      if (d) return d;
    }
    return 0;
  });
}

function keyOf(v) {
  return typeof v === 'string' ? `s:${v.toLowerCase()}` : `${typeof v}:${v}`;
}

// ---------- Building ----------

// Lays out the pivot: returns { rows: [[input]], styles: [{ r, c, style }], widths: { col: px } }.
export function layout(wb, def) {
  const { sheetId, range, rows: rowFields, columns: colFields, values } = resolve(wb, def);
  const vals = values.length ? values : [{ ...(rowFields[0] || colFields[0]), agg: 'count' }];
  const sheetRef = quoteSheetName(wb.sheetName(sheetId));
  const colRange = (col) => `${sheetRef}!${absolute({ r0: range.r0 + 1, c0: col, r1: range.r1, c1: col })}`;

  const rowItems = rowCombinations(wb, sheetId, range, rowFields);
  const colField = colFields[0];
  const colItems = colField ? distinct(wb, sheetId, range, colField.col) : [];
  if (rowItems.length > MAX_ROW_ITEMS) throw new Error(`Too many different row values (${rowItems.length}); pick a field with fewer than ${MAX_ROW_ITEMS}.`);
  if (colItems.length > MAX_COLUMN_ITEMS) throw new Error(`Too many different column values (${colItems.length}); pick a field with fewer than ${MAX_COLUMN_ITEMS}.`);

  const out = [];
  const styles = [];
  const labelCols = Math.max(1, rowFields.length);
  const valueName = (v) => `${AGGREGATES[v.agg].label} of ${v.name}`;
  const numFmtOf = (v) => (v.agg === 'count' ? null : wb.getNumFmt(sheetId, range.r0 + 1, v.col));
  const fmtFor = (field) => wb.getNumFmt(sheetId, range.r0 + 1, field.col);

  // Header row. With a column field, each column value gets one column per value field.
  const header = rowFields.length ? rowFields.map((f) => text(f.name)) : [text('')];
  const dataCols = []; // [{ colItem (or undefined), value }]
  const groups = colField ? [...colItems.map((item) => ({ item })), { item: undefined, total: true }] : [{ item: undefined }];
  for (const group of groups) {
    for (const v of vals) {
      dataCols.push({ ...group, value: v });
      let label;
      if (!colField) label = valueName(v);
      else if (group.total) label = vals.length > 1 ? `Total ${valueName(v)}` : 'Grand Total';
      else label = vals.length > 1 ? `${display(group.item)} - ${valueName(v)}` : display(group.item);
      header.push(text(label));
    }
  }
  let topRows = 0;
  if (colField) {
    // A caption row above the header, like Excel's "Sum of Amount | Quarter ▾".
    const caption = Array(header.length).fill(null);
    caption[0] = text(vals.length === 1 ? valueName(vals[0]) : 'Values');
    caption[labelCols] = text(`${colField.name} →`);
    out.push(caption);
    styles.push({ r: 0, c: 0, style: { italic: true } }, { r: 0, c: labelCols, style: { italic: true } });
    topRows = 1;
  }
  out.push(header);
  for (let c = 0; c < header.length; c++) styles.push({ r: topRows, c, style: { bold: true, fill: HEADER_FILL, border: { bottom: 'thin' } } });

  // One formula per cell: criteria for the row's labels and the column's value.
  const formula = (v, rowValues, rowNumber, colItem) => {
    const criteria = [];
    if (rowValues) rowFields.forEach((f, i) => criteria.push(colRange(f.col), criterion(rowValues[i], `$${cellName(rowNumber, i)}`)));
    if (colItem !== undefined) criteria.push(colRange(colField.col), criterion(colItem, null));
    const agg = AGGREGATES[v.agg];
    const target = colRange(v.col);
    if (v.agg === 'count') {
      if (!criteria.length) return `=COUNTA(${target})`;
      return `=COUNTIFS(${criteria.join(',')},${target},"<>")`;
    }
    if (!criteria.length) return v.agg === 'average' ? `=IFERROR(AVERAGE(${target}),"")` : `=SUM(${target})`;
    if (v.agg === 'average') {
      // Counts numbers only (any number is above -1E+307), as AVERAGEIFS would.
      return `=IFERROR(SUMIFS(${target},${criteria.join(',')})/COUNTIFS(${criteria.join(',')},${target},">-1E+307"),"")`;
    }
    return `=${agg.ifs}(${target},${criteria.join(',')})`;
  };

  const bodyStart = out.length;
  const items = rowFields.length ? rowItems : [null];
  items.forEach((rowValues) => {
    const r = out.length;
    const row = rowValues ? rowValues.map(labelInput) : [text('Total')];
    for (const col of dataCols) row.push(formula(col.value, rowValues, r, col.item));
    out.push(row);
    rowFields.forEach((f, i) => {
      if (rowValues && typeof rowValues[i] === 'number' && fmtFor(f) !== 'General') styles.push({ r, c: i, style: { numFmt: fmtFor(f), align: 'left' } });
    });
  });

  // Grand total row (only meaningful with row fields).
  if (rowFields.length) {
    const r = out.length;
    const row = [text('Grand Total'), ...Array(labelCols - 1).fill(null)];
    for (const col of dataCols) row.push(formula(col.value, null, r, col.item));
    out.push(row);
    for (let c = 0; c < row.length; c++) styles.push({ r, c, style: { bold: true, fill: TOTAL_FILL, border: { top: 'thin' } } });
  }

  // Number formats for the summary cells.
  dataCols.forEach((col, i) => {
    const fmt = numFmtOf(col.value);
    const c = labelCols + i;
    for (let r = bodyStart; r < out.length; r++) {
      if (fmt && fmt !== 'General') styles.push({ r, c, style: { numFmt: fmt } });
      else if (col.value.agg === 'average') styles.push({ r, c, style: { numFmt: '#,##0.00' } });
    }
  });

  const widths = {};
  header.forEach((h, c) => {
    const len = Math.max(String(h).length - 1, ...out.slice(bodyStart, bodyStart + 200).map((row) => (typeof row[c] === 'string' && !row[c].startsWith('=') ? row[c].length - 1 : 6)));
    widths[c] = Math.min(320, Math.max(64, Math.round(len * 7 + 16)));
  });
  return { rows: out, styles, widths };
}

// Writes the pivot onto its sheet (creating it if needed), as one undoable step.
// Returns the sheet id.
export function writePivot(wb, def) {
  const built = layout(wb, def); // throws before anything changes
  let id;
  wb.transact(() => {
    id = wb.sheetIds.find((x) => wb.sheetName(x) === def.sheet);
    if (id === undefined) {
      id = wb.addSheet('PivotTable');
      def = { ...def, sheet: wb.sheetName(id) };
      if (!def.name) def.name = def.sheet;
    } else {
      // Refresh: clear what was there before.
      const used = wb.usedSize(id);
      if (used.rows && used.cols) {
        const all = { r0: 0, c0: 0, r1: used.rows - 1, c1: used.cols - 1 };
        wb.setInputs(id, 0, 0, Array.from({ length: used.rows }, () => Array(used.cols).fill(null)));
        wb.clearStylesIn(id, all);
      }
    }
    const width = Math.max(...built.rows.map((r) => r.length));
    wb.setInputs(
      id,
      0,
      0,
      built.rows.map((r) => Array.from({ length: width }, (_, i) => (r[i] === undefined ? null : r[i]))),
    );
    for (const { r, c, style } of built.styles) wb.setStyleIn(id, { r0: r, c0: c, r1: r, c1: c }, style);
    wb.editMeta(id, 'colWidths', (m) => {
      m.clear();
      for (const [c, px] of Object.entries(built.widths)) if (px !== 64) m.set(Number(c), px);
    });
    wb.editMeta(id, 'frozen', () => ({ rows: def.columns.length ? 2 : 1, cols: 0 }));
    wb.pivots = [...wb.pivots.filter((p) => p.sheet !== def.sheet), { ...def, sheet: wb.sheetName(id) }];
    wb.activeId = id;
  });
  return id;
}

// Re-reads the data (including rows added below it since) and rebuilds the pivot.
export function refreshPivot(wb, def) {
  const sheetId = wb.sheetIds.find((id) => wb.sheetName(id).toLowerCase() === def.source.sheet.toLowerCase());
  let next = def;
  if (sheetId !== undefined) {
    const range = parseRef(def.source.range);
    if (range) next = { ...def, source: { ...def.source, range: rangeName(growDown(wb, sheetId, range)) } };
  }
  return writePivot(wb, next);
}

export function pivotOnSheet(wb, sheetId) {
  return wb.pivots.find((p) => p.sheet === wb.sheetName(sheetId));
}

// The data range grows to take in rows added directly below it.
function growDown(wb, sheetId, range) {
  let r1 = range.r1;
  const filled = (r) => {
    for (let c = range.c0; c <= range.c1; c++) {
      const v = wb.getValue(sheetId, r, c);
      if (v !== null && v !== undefined && v !== '') return true;
    }
    return false;
  };
  while (filled(r1 + 1)) r1++;
  return { ...range, r1 };
}

// ---------- Helpers ----------

function absolute(range) {
  return rangeName(range).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
}

function text(s) {
  return `'${s}`;
}

function display(v) {
  if (v === null) return '(blank)';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function labelInput(v) {
  if (v === null) return text('(blank)');
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  return text(v);
}

// The criteria argument for one value. Plain values point at the label cell
// (so the formula is easy to read); blanks and text that criteria would treat
// as an operator or wildcard are written out exactly.
function criterion(value, labelRef) {
  if (value === null) return '"="';
  if (typeof value === 'string' && (/^[<>=]/.test(value) || /[*?~]/.test(value) || !labelRef)) {
    return `"=${value.replace(/[~*?]/g, '~$&').replace(/"/g, '""')}"`;
  }
  if (!labelRef) return typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : String(value);
  return labelRef;
}
