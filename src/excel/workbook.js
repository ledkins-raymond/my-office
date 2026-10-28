// The spreadsheet's data: cell contents and formulas live in HyperFormula (the
// calculation engine); formatting, column widths, row heights and frozen panes
// live alongside it here. Every change goes through transact() so one Undo
// reverts both halves together.
import { HyperFormula } from 'hyperformula';

export const DEFAULT_COL_PX = 64;
export const DEFAULT_ROW_PX = 20;
const UNDO_LIMIT = 100;

const ENGINE_CONFIG = {
  licenseKey: 'gpl-v3',
  dateFormats: ['MM/DD/YYYY', 'MM/DD/YY', 'YYYY-MM-DD', 'MM/DD'],
  timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
  currencySymbol: ['$'],
  useArrayArithmetic: true,
  maxRows: 1048576,
  maxColumns: 16384,
  // HyperFormula keeps its own undo history; ours decides how far back it goes.
  undoLimit: 100000,
};

// Number formats Excel picks on its own when you type a date, money or a percentage.
const AUTO_FORMATS = {
  NUMBER_DATE: 'm/d/yyyy',
  NUMBER_DATETIME: 'm/d/yyyy h:mm',
  NUMBER_TIME: 'h:mm AM/PM',
  NUMBER_CURRENCY: '$#,##0.00',
  NUMBER_PERCENT: '0%',
};

// ---------- Addresses ----------

export function colName(c) {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

export function cellName(r, c) {
  return `${colName(c)}${r + 1}`;
}

export function rangeName({ r0, c0, r1, c1 }) {
  return r0 === r1 && c0 === c1 ? cellName(r0, c0) : `${cellName(r0, c0)}:${cellName(r1, c1)}`;
}

// "B3" → { r: 2, c: 1 }; "B3:D5" → a range. Returns null for anything else.
export function parseRef(text) {
  const m = /^\s*\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?\s*$/.exec(text);
  if (!m) return null;
  const col = (letters) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const a = { r: Number(m[2]) - 1, c: col(m[1]) };
  const b = m[3] ? { r: Number(m[4]) - 1, c: col(m[3]) } : a;
  if (a.r < 0 || b.r < 0) return null;
  return normRange({ r0: a.r, c0: a.c, r1: b.r, c1: b.c });
}

export function normRange({ r0, c0, r1, c1 }) {
  return { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
}

const key = (r, c) => `${r},${c}`;
const unkey = (k) => k.split(',').map(Number);

export function quoteSheetName(name) {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseRef(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

// ---------- Workbook ----------

export class Workbook {
  constructor(model) {
    this.hf = HyperFormula.buildEmpty(ENGINE_CONFIG);
    // Excel formulas use TRUE and FALSE as plain words; HyperFormula wants TRUE().
    this.hf.addNamedExpression('TRUE', '=TRUE()');
    this.hf.addNamedExpression('FALSE', '=FALSE()');
    this.meta = new Map(); // sheet id → { styles, colWidths, rowHeights, frozen, cached }
    this.order = []; // sheet ids in tab order
    this.undoStack = [];
    this.redoStack = [];
    this.listeners = new Set();
    this.load(model || { sheets: [{ name: 'Sheet1', cells: [] }] });
  }

  onChange(fn) {
    this.listeners.add(fn);
  }

  // kind is 'edit', 'load' or 'view' (switching sheets changes nothing worth saving).
  emit(kind = 'edit') {
    for (const fn of this.listeners) fn(kind);
  }

  load(model) {
    // Pivot table definitions, so they can be refreshed: see pivot.js.
    this.pivots = (model.pivots || []).map((p) => ({ ...p }));
    this.hf.suspendEvaluation();
    for (const id of this.order) this.hf.removeSheet(id);
    this.meta.clear();
    this.order = [];
    for (const sheet of model.sheets) {
      const id = this.hf.getSheetId(this.hf.addSheet(uniqueName(this.hf, sheet.name)));
      const meta = newMeta();
      meta.frozen = { rows: sheet.frozen?.rows || 0, cols: sheet.frozen?.cols || 0 };
      for (const [c, px] of Object.entries(sheet.colWidths || {})) meta.colWidths.set(Number(c), px);
      for (const [r, px] of Object.entries(sheet.rowHeights || {})) meta.rowHeights.set(Number(r), px);
      const rows = [];
      for (const cell of sheet.cells || []) {
        if (cell.style) meta.styles.set(key(cell.r, cell.c), cell.style);
        if (cell.input === undefined || cell.input === null || cell.input === '') continue;
        (rows[cell.r] ||= [])[cell.c] = cell.input;
        if (cell.value !== undefined) meta.cached.set(key(cell.r, cell.c), cell.value);
      }
      if (rows.length) this.hf.setSheetContent(id, densify(rows));
      this.meta.set(id, meta);
      this.order.push(id);
    }
    for (const { name, ref } of model.names || []) {
      try {
        this.hf.addNamedExpression(name, `=${ref.replace(/^=/, '')}`);
      } catch {
        // A name HyperFormula can't use (for example one that clashes with a cell address) is skipped.
      }
    }
    this.hf.resumeEvaluation();
    this.hf.clearUndoStack();
    this.hf.clearRedoStack();
    this.undoStack = [];
    this.redoStack = [];
    this.activeId = this.order[Math.min(model.activeSheet || 0, this.order.length - 1)];
    this.emit('load');
  }

  // ---------- Reading ----------

  get sheetIds() {
    return this.order;
  }

  sheetName(id) {
    return this.hf.getSheetName(id);
  }

  metaOf(id) {
    return this.meta.get(id);
  }

  // What the cell shows in the formula bar: the formula, or the typed value.
  getInput(id, r, c) {
    const raw = this.hf.getCellSerialized({ sheet: id, row: r, col: c });
    if (raw === null || raw === undefined) return '';
    if (typeof raw === 'string') return raw.startsWith("'") ? raw.slice(1) : raw;
    if (typeof raw === 'boolean') return raw ? 'TRUE' : 'FALSE';
    return String(raw);
  }

  isFormula(id, r, c) {
    return this.hf.doesCellHaveFormula({ sheet: id, row: r, col: c });
  }

  // The calculated value: a number, string, boolean, null, or { error: '#DIV/0!' }.
  getValue(id, r, c) {
    const value = this.hf.getCellValue({ sheet: id, row: r, col: c });
    if (value && typeof value === 'object') {
      // A formula this engine can't run (a function it lacks, say) keeps the
      // result it had when the file was last saved by another program.
      const cached = this.meta.get(id).cached.get(key(r, c));
      if (cached !== undefined && (value.type === 'NAME' || value.type === 'ERROR')) return cached;
      return { error: value.value, message: value.message };
    }
    return value;
  }

  getStyle(id, r, c) {
    return this.meta.get(id).styles.get(key(r, c));
  }

  // The number format to display with: the cell's own, or the one Excel would pick.
  getNumFmt(id, r, c) {
    const own = this.getStyle(id, r, c)?.numFmt;
    if (own) return own;
    return AUTO_FORMATS[this.hf.getCellValueDetailedType({ sheet: id, row: r, col: c })] || 'General';
  }

  colWidth(id, c) {
    return this.meta.get(id).colWidths.get(c) ?? DEFAULT_COL_PX;
  }

  rowHeight(id, r) {
    return this.meta.get(id).rowHeights.get(r) ?? DEFAULT_ROW_PX;
  }

  // Rows and columns in use, counting formatting: { rows, cols } (both 0 when empty).
  usedSize(id) {
    const dims = this.hf.getSheetDimensions(id);
    let rows = dims.height;
    let cols = dims.width;
    const meta = this.meta.get(id);
    for (const k of meta.styles.keys()) {
      const [r, c] = unkey(k);
      rows = Math.max(rows, r + 1);
      cols = Math.max(cols, c + 1);
    }
    for (const c of meta.colWidths.keys()) cols = Math.max(cols, c + 1);
    for (const r of meta.rowHeights.keys()) rows = Math.max(rows, r + 1);
    return { rows, cols };
  }

  // ---------- Changing (everything below is undoable) ----------

  // Runs `fn` as one undoable step. `fn` changes the engine only through
  // this.engine() and this.paste(), so we know how many engine steps to undo.
  transact(fn) {
    if (this.pending) return fn(); // already inside a step: becomes part of it
    const entry = { steps: 0, before: this.snapshot() };
    this.pending = entry;
    try {
      fn();
    } finally {
      this.pending = null;
    }
    entry.after = this.snapshot();
    this.undoStack.push(entry);
    if (this.undoStack.length > UNDO_LIMIT) this.undoStack.shift();
    this.redoStack = [];
    this.emit();
  }

  engine(fn) {
    this.pending.steps++;
    return this.hf.batch(fn);
  }

  // Copies `src` (on sheet `from`) to the top-left of `dest`, adjusting relative references like Excel.
  pasteCells(from, src, to, dest, { cut = false } = {}) {
    const range = { start: { sheet: from, row: src.r0, col: src.c0 }, end: { sheet: from, row: src.r1, col: src.c1 } };
    if (cut) this.hf.cut(range);
    else this.hf.copy(range);
    this.pending.steps++;
    this.hf.paste({ sheet: to, row: dest.r, col: dest.c });
  }

  snapshot() {
    const metas = new Map();
    for (const [id, m] of this.meta) {
      metas.set(id, {
        styles: m.styles,
        colWidths: m.colWidths,
        rowHeights: m.rowHeights,
        frozen: m.frozen,
        cached: m.cached,
      });
    }
    return { metas, order: [...this.order], activeId: this.activeId, pivots: this.pivots };
  }

  restore(snap) {
    for (const [id, m] of snap.metas) this.meta.set(id, { ...m });
    this.order = [...snap.order];
    this.pivots = snap.pivots;
    this.activeId = snap.activeId;
  }

  // Metadata maps are replaced, never edited in place, so snapshots stay intact.
  editMeta(id, field, fn) {
    const meta = this.meta.get(id);
    const copy = field === 'frozen' ? { ...meta.frozen } : new Map(meta[field]);
    meta[field] = fn(copy) || copy;
  }

  undo() {
    const entry = this.undoStack.pop();
    if (!entry) return false;
    for (let i = 0; i < entry.steps; i++) this.hf.undo();
    this.restore(entry.before);
    this.redoStack.push(entry);
    this.emit();
    return true;
  }

  redo() {
    const entry = this.redoStack.pop();
    if (!entry) return false;
    for (let i = 0; i < entry.steps; i++) this.hf.redo();
    this.restore(entry.after);
    this.undoStack.push(entry);
    this.emit();
    return true;
  }

  // Sets what the user typed into one cell, picking a number format like Excel does.
  enter(id, r, c, text) {
    this.transact(() => {
      this.setInputs(id, r, c, [[toInput(text)]]);
      const type = this.hf.getCellValueDetailedType({ sheet: id, row: r, col: c });
      const isFormula = this.isFormula(id, r, c);
      if (!isFormula && AUTO_FORMATS[type] && !this.getStyle(id, r, c)?.numFmt) {
        this.setStyleIn(id, { r0: r, c0: c, r1: r, c1: c }, { numFmt: AUTO_FORMATS[type] });
      }
    });
  }

  // Writes a block of inputs (inside a transaction).
  setInputs(id, r, c, rows) {
    this.engine(() => this.hf.setCellContents({ sheet: id, row: r, col: c }, rows));
    const height = rows.length;
    const width = Math.max(0, ...rows.map((row) => row.length));
    this.editMeta(id, 'cached', (m) => {
      for (const k of [...m.keys()]) {
        const [kr, kc] = unkey(k);
        if (kr >= r && kr < r + height && kc >= c && kc < c + width) m.delete(k);
      }
    });
  }

  setStyleIn(id, range, patch) {
    this.editMeta(id, 'styles', (styles) => {
      for (let r = range.r0; r <= range.r1; r++) {
        for (let c = range.c0; c <= range.c1; c++) {
          const k = key(r, c);
          const next = { ...styles.get(k) };
          for (const [name, value] of Object.entries(typeof patch === 'function' ? patch(next, r, c) : patch)) {
            if (value === null || value === undefined || value === false) delete next[name];
            else next[name] = value;
          }
          if (Object.keys(next).length) styles.set(k, next);
          else styles.delete(k);
        }
      }
    });
  }

  setStyle(id, range, patch) {
    this.transact(() => this.setStyleIn(id, range, patch));
  }

  clearStylesIn(id, range) {
    this.editMeta(id, 'styles', (styles) => {
      for (const k of [...styles.keys()]) {
        const [r, c] = unkey(k);
        if (inRange(range, r, c)) styles.delete(k);
      }
    });
  }

  // what: 'contents' (Delete key), 'formats', or 'all'.
  clear(id, range, what = 'contents') {
    this.transact(() => {
      if (what !== 'formats') {
        const blank = Array.from({ length: range.r1 - range.r0 + 1 }, () => Array(range.c1 - range.c0 + 1).fill(null));
        this.setInputs(id, range.r0, range.c0, blank);
      }
      if (what !== 'contents') this.clearStylesIn(id, range);
    });
  }

  setColWidth(id, cols, px) {
    this.transact(() =>
      this.editMeta(id, 'colWidths', (m) => {
        for (const c of cols) {
          if (px === DEFAULT_COL_PX) m.delete(c);
          else m.set(c, Math.max(0, Math.round(px)));
        }
      }),
    );
  }

  setRowHeight(id, rows, px) {
    this.transact(() =>
      this.editMeta(id, 'rowHeights', (m) => {
        for (const r of rows) {
          if (px === DEFAULT_ROW_PX) m.delete(r);
          else m.set(r, Math.max(0, Math.round(px)));
        }
      }),
    );
  }

  setFrozen(id, rows, cols) {
    this.transact(() => this.editMeta(id, 'frozen', () => ({ rows, cols })));
  }

  // axis 'row' or 'col'; count > 0 inserts before `at`, count < 0 deletes from `at`.
  shift(id, axis, at, count) {
    this.transact(() => {
      this.engine(() => {
        if (axis === 'row') {
          if (count > 0) this.hf.addRows(id, [at, count]);
          else this.hf.removeRows(id, [at, -count]);
        } else if (count > 0) this.hf.addColumns(id, [at, count]);
        else this.hf.removeColumns(id, [at, -count]);
      });
      const moveKey = (r, c) => {
        const pos = axis === 'row' ? r : c;
        if (count < 0 && pos >= at && pos < at - count) return null;
        const moved = pos >= at ? pos + count : pos;
        return axis === 'row' ? key(moved, c) : key(r, moved);
      };
      for (const field of ['styles', 'cached']) {
        this.editMeta(id, field, (m) => {
          const out = new Map();
          for (const [k, v] of m) {
            const nk = moveKey(...unkey(k));
            if (nk) out.set(nk, v);
          }
          return out;
        });
      }
      this.editMeta(id, axis === 'row' ? 'rowHeights' : 'colWidths', (m) => {
        const out = new Map();
        for (const [pos, v] of m) {
          if (count < 0 && pos >= at && pos < at - count) continue;
          out.set(pos >= at ? pos + count : pos, v);
        }
        return out;
      });
    });
  }

  // Copy or move cells, with their formatting (Excel's paste).
  paste(clip, to, dest) {
    const { sheet: from, range: src, cut } = clip;
    const h = src.r1 - src.r0;
    const w = src.c1 - src.c0;
    const target = { r0: dest.r, c0: dest.c, r1: dest.r + h, c1: dest.c + w };
    const styles = this.copyStyles(from, src);
    this.transact(() => {
      this.pasteCells(from, src, to, dest, { cut });
      if (cut) this.clearStylesIn(from, src);
      this.clearStylesIn(to, target);
      this.applyStyles(to, target, styles);
    });
    return target;
  }

  copyStyles(id, src) {
    const out = [];
    for (let r = src.r0; r <= src.r1; r++) {
      const row = [];
      for (let c = src.c0; c <= src.c1; c++) row.push(this.getStyle(id, r, c));
      out.push(row);
    }
    return out;
  }

  applyStyles(id, target, styles) {
    this.editMeta(id, 'styles', (m) => {
      styles.forEach((row, i) =>
        row.forEach((style, j) => {
          if (style) m.set(key(target.r0 + i, target.c0 + j), style);
        }),
      );
    });
  }

  // Pastes plain text rows (from another program) as typed values.
  pasteValues(id, dest, rows) {
    this.transact(() => this.setInputs(id, dest.r, dest.c, rows.map((row) => row.map(toInput))));
    return { r0: dest.r, c0: dest.c, r1: dest.r + rows.length - 1, c1: dest.c + Math.max(...rows.map((r) => r.length)) - 1 };
  }

  // Extends `src` to fill `target` (which contains it), like dragging the fill handle,
  // Ctrl+D or Ctrl+R. Numbers in a steady series continue; everything else repeats.
  fill(id, src, target) {
    this.transact(() => {
      const down = target.r1 > src.r1 || target.r0 < src.r0;
      const lines = down ? [src.c0, src.c1] : [src.r0, src.r1];
      const srcLen = down ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1;
      const start = down ? target.r0 : target.c0;
      const end = down ? target.r1 : target.c1;
      const srcStart = down ? src.r0 : src.c0;
      const styles = this.copyStyles(id, src);

      // Copy the source block repeatedly; formulas adjust their references.
      for (let pos = start; pos <= end; pos++) {
        const offset = pos - srcStart;
        if (offset >= 0 && offset < srcLen) continue;
        const from = srcStart + (((offset % srcLen) + srcLen) % srcLen);
        const one = down ? { r0: from, r1: from, c0: src.c0, c1: src.c1 } : { r0: src.r0, r1: src.r1, c0: from, c1: from };
        this.pasteCells(id, one, id, down ? { r: pos, c: src.c0 } : { r: src.r0, c: pos });
        const styleLine = down ? [styles[from - src.r0]] : styles.map((row) => [row[from - src.c0]]);
        this.applyStyles(id, down ? { r0: pos, c0: src.c0 } : { r0: src.r0, c0: pos }, styleLine);
      }

      // Number series: 1, 2 → 3, 4 …  (a single number just repeats, as in Excel)
      if (srcLen < 2) return;
      for (let line = lines[0]; line <= lines[1]; line++) {
        const at = (i) => (down ? { sheet: id, row: srcStart + i, col: line } : { sheet: id, row: line, col: srcStart + i });
        const values = [];
        for (let i = 0; i < srcLen; i++) {
          if (this.hf.doesCellHaveFormula(at(i))) break;
          const v = this.hf.getCellValue(at(i));
          if (typeof v !== 'number') break;
          values.push(v);
        }
        if (values.length !== srcLen) continue;
        const step = (values[srcLen - 1] - values[0]) / (srcLen - 1);
        const updates = [];
        for (let pos = start; pos <= end; pos++) {
          const offset = pos - srcStart;
          if (offset >= 0 && offset < srcLen) continue;
          updates.push([pos, values[0] + step * offset]);
        }
        for (const [pos, v] of updates) {
          const value = Math.round(v * 1e10) / 1e10;
          this.setInputs(id, down ? pos : line, down ? line : pos, [[value]]);
        }
      }
    });
  }

  // Sorts the rows of `range` by column `byCol`. Text after numbers, blanks last, as in Excel.
  sort(id, range, byCol, ascending = true) {
    const rows = [];
    for (let r = range.r0; r <= range.r1; r++) {
      const inputs = [];
      const styles = [];
      for (let c = range.c0; c <= range.c1; c++) {
        inputs.push(this.hf.getCellSerialized({ sheet: id, row: r, col: c }) ?? null);
        styles.push(this.getStyle(id, r, c));
      }
      rows.push({ sortKey: this.hf.getCellValue({ sheet: id, row: r, col: byCol }), inputs, styles });
    }
    const rank = (v) => (v === null || v === '' ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
    rows.sort((a, b) => {
      const ra = rank(a.sortKey);
      const rb = rank(b.sortKey);
      if (ra !== rb || ra >= 2) return ra - rb; // blanks and errors stay at the bottom either way
      const cmp =
        ra === 0 ? a.sortKey - b.sortKey : String(a.sortKey).localeCompare(String(b.sortKey), undefined, { sensitivity: 'base' });
      return ascending ? cmp : -cmp;
    });
    this.transact(() => {
      this.setInputs(
        id,
        range.r0,
        range.c0,
        rows.map((row) => row.inputs),
      );
      this.clearStylesIn(id, range);
      this.applyStyles(id, range, rows.map((row) => row.styles));
    });
  }

  // Adds a sheet named Sheet2, Sheet3… (or PivotTable1… for another base name) and returns its id.
  addSheet(base = 'Sheet') {
    let n = base === 'Sheet' ? this.order.length + 1 : 1;
    while (this.hf.doesSheetExist(`${base}${n}`)) n++;
    let id;
    this.transact(() => {
      this.engine(() => {
        id = this.hf.getSheetId(this.hf.addSheet(`${base}${n}`));
      });
      this.meta.set(id, newMeta());
      this.order = [...this.order, id];
      this.activeId = id;
    });
    return id;
  }

  removeSheet(id) {
    if (this.order.length < 2) return;
    this.transact(() => {
      const index = this.order.indexOf(id);
      const name = this.sheetName(id);
      this.engine(() => this.hf.removeSheet(id));
      this.pivots = this.pivots.filter((p) => p.sheet !== name);
      this.order = this.order.filter((x) => x !== id);
      this.activeId = this.order[Math.max(0, index - 1)];
    });
  }

  // Returns an error message, or null when renamed.
  renameSheet(id, name) {
    name = name.trim();
    if (!name) return 'A sheet name cannot be blank.';
    if (name.length > 31) return 'A sheet name can be at most 31 characters.';
    if (/[[\]:*?/\\]/.test(name) || /^'|'$/.test(name)) return "A sheet name cannot contain [ ] : * ? / \\ or start or end with '.";
    const existing = this.hf.getSheetId(name);
    if (existing !== undefined && existing !== id) return 'That name is already used by another sheet.';
    if (name === this.sheetName(id)) return null;
    const old = this.sheetName(id);
    this.transact(() => {
      this.engine(() => this.hf.renameSheet(id, name));
      // Pivot tables refer to sheets by name.
      this.pivots = this.pivots.map((p) => ({
        ...p,
        sheet: p.sheet === old ? name : p.sheet,
        source: p.source.sheet === old ? { ...p.source, sheet: name } : p.source,
      }));
    });
    return null;
  }

  moveSheet(id, toIndex) {
    this.transact(() => {
      const order = this.order.filter((x) => x !== id);
      order.splice(toIndex, 0, id);
      this.order = order;
    });
  }

  setActive(id) {
    this.activeId = id;
    this.emit('view');
  }

  // ---------- Saving ----------

  toModel() {
    const sheets = this.order.map((id) => {
      const meta = this.meta.get(id);
      const cells = [];
      const serialized = this.hf.getSheetSerialized(id);
      const seen = new Set();
      serialized.forEach((row, r) =>
        row.forEach((input, c) => {
          if (input === null || input === undefined || input === '') return;
          const cell = { r, c, input };
          if (typeof input === 'string' && input.startsWith('=')) {
            const v = this.getValue(id, r, c);
            if (v !== null && v !== undefined) cell.value = v && typeof v === 'object' ? { error: v.error } : v;
          } else if (typeof input === 'string' && !input.startsWith("'")) {
            // Typed values such as "1/5/2026" or "$12" are saved as the numbers they stand for.
            const v = this.hf.getCellValue({ sheet: id, row: r, col: c });
            cell.input = v && typeof v === 'object' ? `'${input}` : typeof v === 'string' ? `'${v}` : v;
          }
          const style = this.cellStyleForSave(id, r, c);
          if (style) cell.style = style;
          cells.push(cell);
          seen.add(key(r, c));
        }),
      );
      for (const [k, style] of meta.styles) {
        if (seen.has(k)) continue;
        const [r, c] = unkey(k);
        cells.push({ r, c, style });
      }
      return {
        name: this.sheetName(id),
        cells,
        colWidths: Object.fromEntries(meta.colWidths),
        rowHeights: Object.fromEntries(meta.rowHeights),
        frozen: { ...meta.frozen },
      };
    });
    const names = this.hf
      .listNamedExpressions()
      .filter((name) => name !== 'TRUE' && name !== 'FALSE')
      .map((name) => ({ name, ref: this.hf.getNamedExpressionFormula(name).replace(/^=/, '') }));
    return { activeSheet: Math.max(0, this.order.indexOf(this.activeId)), names, sheets, pivots: this.pivots.map((p) => ({ ...p })) };
  }

  cellStyleForSave(id, r, c) {
    const style = this.getStyle(id, r, c);
    const numFmt = this.getNumFmt(id, r, c);
    if (numFmt === 'General' || style?.numFmt) return style;
    return { ...style, numFmt };
  }
}

// ---------- Helpers ----------

function newMeta() {
  return {
    styles: new Map(),
    colWidths: new Map(),
    rowHeights: new Map(),
    frozen: { rows: 0, cols: 0 },
    cached: new Map(),
  };
}

function inRange(range, r, c) {
  return r >= range.r0 && r <= range.r1 && c >= range.c0 && c <= range.c1;
}

function densify(rows) {
  const width = Math.max(0, ...rows.map((row) => (row ? row.length : 0)));
  return Array.from(rows, (row) => Array.from({ length: width }, (_, c) => (row && row[c] !== undefined ? row[c] : null)));
}

function uniqueName(hf, name) {
  let candidate = name || 'Sheet';
  for (let n = 2; hf.doesSheetExist(candidate); n++) candidate = `${name} (${n})`;
  return candidate;
}

// Turns what was typed into engine input. A leading apostrophe forces text, as in Excel.
export function toInput(text) {
  if (text === null || text === undefined) return null;
  const s = String(text);
  if (s === '') return null;
  if (s.startsWith("'")) return s;
  if (s.startsWith('=')) return s.length > 1 ? s : "'=";
  const upper = s.trim().toUpperCase();
  if (upper === 'TRUE') return true;
  if (upper === 'FALSE') return false;
  // Thousands separators: "1,234.5" → 1234.5
  if (/^[-+]?\$?\d{1,3}(,\d{3})+(\.\d+)?%?$/.test(s.trim())) return s.trim().replace(/,/g, '');
  return s;
}
