// The spreadsheet grid: drawn on a canvas over a scrolling area that provides
// the scrollbars. Handles selection, typing into cells, pointing at cells
// while writing a formula, the fill handle, resizing columns and rows, and
// frozen panes.
import { cellName, normRange, colName, rangeName, parseRef } from './workbook.js';
import { formatValue, fitNumber, isError, editText } from './format.js';

const HEADER_H = 22;
const MAX_ROWS = 1048576;
const MAX_COLS = 16384;
const MIN_ROWS = 200;
const MIN_COLS = 40;
const DEFAULT_FONT = 'Calibri';
const FONT_FALLBACK = "Carlito, 'Segoe UI', Arial, sans-serif";
const ACCENT = '#217346';
const GRID_LINE = '#e1e1e1';
const REF_COLORS = ['#3b6fd8', '#c0392b', '#7b3fb0', '#1e8449', '#b9770e', '#0e6f80'];
const POINT_AFTER = /(^=|[=(,+\-*/^&<>:;%\s])$/;

// Sizes along one axis (rows or columns), with fast position lookups.
class Axis {
  constructor(sizeOf, min, max) {
    this.sizeOf = sizeOf;
    this.min = min;
    this.max = max;
    this.count = min;
    this.invalidate();
  }

  invalidate() {
    this.prefix = null;
  }

  ensure(count) {
    const next = Math.min(this.max, Math.max(this.count, count));
    if (next !== this.count) {
      this.count = next;
      this.invalidate();
    }
  }

  build() {
    if (this.prefix) return;
    const p = new Float64Array(this.count + 1);
    for (let i = 0; i < this.count; i++) p[i + 1] = p[i] + this.sizeOf(i);
    this.prefix = p;
  }

  pos(i) {
    this.build();
    return this.prefix[Math.min(Math.max(i, 0), this.count)];
  }

  size(i) {
    return this.sizeOf(i);
  }

  get total() {
    return this.pos(this.count);
  }

  // Index of the row/column at offset x (clamped to the axis).
  at(x) {
    this.build();
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.prefix[mid] <= x) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}

export class Grid {
  constructor(root, wb, hooks = {}) {
    this.root = root;
    this.wb = wb;
    this.hooks = hooks;
    this.sel = { r0: 0, c0: 0, r1: 0, c1: 0 };
    this.active = { r: 0, c: 0 };
    this.anchor = { r: 0, c: 0 };
    this.edit = null; // { mode: 'enter' | 'edit', original, point }
    this.clip = null; // copied/cut range shown with a dashed border
    this.fillTarget = null;
    this.views = new Map(); // per-sheet selection and scroll position

    root.innerHTML = '';
    this.scroller = el('div', 'grid-scroller');
    this.sizer = el('div', 'grid-sizer');
    this.canvas = el('canvas', 'grid-canvas');
    this.input = el('textarea', 'cell-editor');
    this.input.spellcheck = false;
    this.input.setAttribute('aria-label', 'Cell contents');
    this.scroller.append(this.sizer);
    root.append(this.scroller, this.canvas, this.input);
    this.ctx = this.canvas.getContext('2d');

    this.rows = new Axis((r) => this.wb.rowHeight(this.sheet, r), MIN_ROWS, MAX_ROWS);
    this.cols = new Axis((c) => this.wb.colWidth(this.sheet, c), MIN_COLS, MAX_COLS);

    this.bindEvents();
    wb.onChange(() => this.refresh());
  }

  // Call once the page has its hooks ready.
  start() {
    new ResizeObserver(() => this.layout()).observe(this.root);
    this.refresh();
  }

  // ---------- Sheet and geometry ----------

  get sheet() {
    return this.wb.activeId;
  }

  get frozen() {
    return this.wb.metaOf(this.sheet).frozen;
  }

  refresh() {
    if (this.shownSheet !== this.sheet) this.switchSheet();
    const used = this.wb.usedSize(this.sheet);
    this.rows.ensure(used.rows + 50);
    this.cols.ensure(used.cols + 10);
    this.rows.invalidate();
    this.cols.invalidate();
    this.layout();
  }

  switchSheet() {
    if (this.shownSheet !== undefined && this.wb.metaOf(this.shownSheet)) {
      this.views.set(this.shownSheet, {
        sel: this.sel,
        active: this.active,
        top: this.scroller.scrollTop,
        left: this.scroller.scrollLeft,
      });
    }
    this.shownSheet = this.sheet;
    this.rows.count = MIN_ROWS;
    this.cols.count = MIN_COLS;
    const view = this.views.get(this.sheet);
    this.sel = view?.sel || { r0: 0, c0: 0, r1: 0, c1: 0 };
    this.active = view?.active || { r: 0, c: 0 };
    this.anchor = { ...this.active };
    const used = this.wb.usedSize(this.sheet);
    this.rows.ensure(Math.max(used.rows, this.sel.r1) + 50);
    this.cols.ensure(Math.max(used.cols, this.sel.c1) + 10);
    this.rows.invalidate();
    this.cols.invalidate();
    this.updateSizer();
    this.scroller.scrollTop = view?.top || 0;
    this.scroller.scrollLeft = view?.left || 0;
    this.hooks.onSelection?.();
  }

  get headerW() {
    const digits = String(Math.min(this.rows.count, this.lastVisibleRow() + 1)).length;
    return Math.max(36, digits * 8 + 14);
  }

  updateSizer() {
    this.sizer.style.width = `${this.headerW + this.cols.total}px`;
    this.sizer.style.height = `${HEADER_H + this.rows.total}px`;
  }

  layout() {
    const w = this.scroller.clientWidth;
    const h = this.scroller.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    if (this.canvas.width !== Math.round(w * dpr) || this.canvas.height !== Math.round(h * dpr)) {
      this.canvas.width = Math.round(w * dpr);
      this.canvas.height = Math.round(h * dpr);
      this.canvas.style.width = `${w}px`;
      this.canvas.style.height = `${h}px`;
    }
    this.width = w;
    this.height = h;
    this.updateSizer();
    this.draw();
    this.placeEditor();
  }

  frozenW() {
    return this.cols.pos(this.frozen.cols);
  }

  frozenH() {
    return this.rows.pos(this.frozen.rows);
  }

  // Screen position of a column's left edge / a row's top edge.
  colX(c) {
    const base = this.headerW + this.cols.pos(c);
    return c < this.frozen.cols ? base : base - this.scroller.scrollLeft;
  }

  rowY(r) {
    const base = HEADER_H + this.rows.pos(r);
    return r < this.frozen.rows ? base : base - this.scroller.scrollTop;
  }

  cellRect(r, c) {
    return { x: this.colX(c), y: this.rowY(r), w: this.cols.size(c), h: this.rows.size(r) };
  }

  colAtX(x) {
    const fw = this.frozenW();
    const local = x - this.headerW;
    if (local < fw) return this.cols.at(Math.max(0, local));
    return Math.max(this.frozen.cols, this.cols.at(local + this.scroller.scrollLeft));
  }

  rowAtY(y) {
    const fh = this.frozenH();
    const local = y - HEADER_H;
    if (local < fh) return this.rows.at(Math.max(0, local));
    return Math.max(this.frozen.rows, this.rows.at(local + this.scroller.scrollTop));
  }

  firstScrollRow() {
    return Math.max(this.frozen.rows, this.rows.at(this.scroller.scrollTop + this.frozenH()));
  }

  lastVisibleRow() {
    return this.rows.at(this.scroller.scrollTop + (this.height || 800) - HEADER_H);
  }

  firstScrollCol() {
    return Math.max(this.frozen.cols, this.cols.at(this.scroller.scrollLeft + this.frozenW()));
  }

  lastVisibleCol() {
    return this.cols.at(this.scroller.scrollLeft + (this.width || 1200) - this.headerW);
  }

  // ---------- Drawing ----------

  draw() {
    if (!this.width) return;
    const ctx = this.ctx;
    const dpr = window.devicePixelRatio || 1;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, this.width, this.height);

    const hw = this.headerW;
    const fw = this.frozenW();
    const fh = this.frozenH();
    const fr = this.frozen.rows;
    const fc = this.frozen.cols;
    const r0 = this.firstScrollRow();
    const r1 = this.lastVisibleRow();
    const c0 = this.firstScrollCol();
    const c1 = this.lastVisibleCol();
    this.visible = { r0, r1, c0, c1 };

    const right = this.width;
    const bottom = this.height;
    const regions = [
      [r0, r1, c0, c1, hw + fw, HEADER_H + fh, right, bottom],
      [0, fr - 1, c0, c1, hw + fw, HEADER_H, right, HEADER_H + fh],
      [r0, r1, 0, fc - 1, hw, HEADER_H + fh, hw + fw, bottom],
      [0, fr - 1, 0, fc - 1, hw, HEADER_H, hw + fw, HEADER_H + fh],
    ];
    for (const [ra, rb, ca, cb, x0, y0, x1, y1] of regions) {
      if (ra > rb || ca > cb || x1 <= x0 || y1 <= y0) continue;
      ctx.save();
      ctx.beginPath();
      ctx.rect(x0, y0, x1 - x0, y1 - y0);
      ctx.clip();
      this.drawCells(ra, rb, ca, cb);
      this.drawSelection();
      ctx.restore();
    }

    if (fr) line(ctx, hw, HEADER_H + fh - 0.5, right, HEADER_H + fh - 0.5, '#9a9a9a', 1);
    if (fc) line(ctx, hw + fw - 0.5, HEADER_H, hw + fw - 0.5, bottom, '#9a9a9a', 1);
    this.drawHeaders(r0, r1, c0, c1);
  }

  drawCells(r0, r1, c0, c1) {
    const ctx = this.ctx;
    const wb = this.wb;
    const id = this.sheet;

    // Fills
    for (let r = r0; r <= r1; r++) {
      if (!this.rows.size(r)) continue;
      for (let c = c0; c <= c1; c++) {
        const fill = wb.getStyle(id, r, c)?.fill;
        if (!fill) continue;
        const { x, y, w, h } = this.cellRect(r, c);
        ctx.fillStyle = fill;
        ctx.fillRect(x, y, w, h);
      }
    }

    // Gridlines (skipped under filled cells, as in Excel)
    ctx.beginPath();
    ctx.strokeStyle = GRID_LINE;
    ctx.lineWidth = 1;
    const top = this.rowY(r0);
    const bottomY = this.rowY(r1) + this.rows.size(r1);
    const left = this.colX(c0);
    const rightX = this.colX(c1) + this.cols.size(c1);
    for (let c = c0; c <= c1; c++) {
      const x = Math.round(this.colX(c) + this.cols.size(c)) - 0.5;
      ctx.moveTo(x, top);
      ctx.lineTo(x, bottomY);
    }
    for (let r = r0; r <= r1; r++) {
      const y = Math.round(this.rowY(r) + this.rows.size(r)) - 0.5;
      ctx.moveTo(left, y);
      ctx.lineTo(rightX, y);
    }
    ctx.stroke();

    // Text. Start a few columns early so long text from off-screen cells still spills in.
    const textStart = Math.max(0, c0 - 12);
    for (let r = r0; r <= r1; r++) {
      if (!this.rows.size(r)) continue;
      for (let c = textStart; c <= c1; c++) this.drawText(r, c, c0, c1);
    }

    // Borders
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const border = wb.getStyle(id, r, c)?.border;
        if (border) this.drawBorder(r, c, border);
      }
    }
  }

  fontFor(style) {
    const size = ((style?.size || 11) * 4) / 3;
    const family = style?.font ? `'${style.font}', ${FONT_FALLBACK}` : `${DEFAULT_FONT}, ${FONT_FALLBACK}`;
    return `${style?.italic ? 'italic ' : ''}${style?.bold ? 'bold ' : ''}${size}px ${family}`;
  }

  drawText(r, c, c0, c1) {
    const wb = this.wb;
    const id = this.sheet;
    const value = wb.getValue(id, r, c);
    if (value === null || value === undefined || value === '') return;
    const style = wb.getStyle(id, r, c) || {};
    const numFmt = wb.getNumFmt(id, r, c);
    const ctx = this.ctx;
    ctx.font = this.fontFor(style);
    const measure = (t) => ctx.measureText(t).width;
    const rect = this.cellRect(r, c);
    if (!rect.w) return;
    const pad = 3;
    const isNum = typeof value === 'number';
    const align = style.align || (isNum ? 'right' : typeof value === 'boolean' || isError(value) ? 'center' : 'left');

    let text = isNum ? fitNumber(value, numFmt, rect.w - pad * 2, measure) : formatValue(value, numFmt);
    // Accounting-style padding codes ("_(") leave spaces SSF renders literally; trim for drawing.
    if (isNum) text = text.replace(/\s+$/, '');

    // Where the text may spill: across empty neighbours, for unwrapped text only.
    let clipX0 = rect.x;
    let clipX1 = rect.x + rect.w;
    if (!isNum && !style.wrap) {
      const width = measure(text) + pad * 2;
      if (width > rect.w) {
        const spillRight = align === 'left' || align === 'center';
        const spillLeft = align === 'right' || align === 'center';
        if (spillRight) {
          for (let n = c + 1; n <= c1 + 1 && clipX1 - rect.x < width; n++) {
            if (this.hasContent(r, n)) break;
            clipX1 = this.colX(n) + this.cols.size(n);
          }
        }
        if (spillLeft) {
          for (let n = c - 1; n >= 0 && rect.x + rect.w - clipX0 < width; n--) {
            if (this.hasContent(r, n)) break;
            clipX0 = this.colX(n);
          }
        }
      }
    }
    if (clipX1 < this.colX(c0) && c < c0) return; // entirely off-screen

    const lines = style.wrap ? wrapLines(text, rect.w - pad * 2, measure) : [text];
    const lineH = (((style.size || 11) * 4) / 3) * 1.2;
    const blockH = lineH * lines.length;
    let y;
    if (style.valign === 'top') y = rect.y + 2;
    else if (style.valign === 'middle') y = rect.y + (rect.h - blockH) / 2;
    else y = rect.y + rect.h - blockH - 2;

    ctx.save();
    ctx.beginPath();
    ctx.rect(clipX0 + 1, rect.y, clipX1 - clipX0 - 2, rect.h - 1);
    ctx.clip();
    const color = isError(value) && !style.color ? '#1f1f1f' : style.color || '#000000';
    ctx.fillStyle = color;
    ctx.textBaseline = 'middle';
    for (const lineText of lines) {
      const tw = measure(lineText);
      let x;
      if (align === 'right') x = rect.x + rect.w - pad - tw;
      else if (align === 'center') x = rect.x + (rect.w - tw) / 2;
      else x = rect.x + pad;
      const cy = y + lineH / 2;
      ctx.fillText(lineText, x, cy);
      if (style.underline) line(ctx, x, Math.round(cy + lineH * 0.38) + 0.5, x + tw, Math.round(cy + lineH * 0.38) + 0.5, color, 1);
      if (style.strike) line(ctx, x, Math.round(cy) + 0.5, x + tw, Math.round(cy) + 0.5, color, 1);
      y += lineH;
    }
    ctx.restore();
  }

  hasContent(r, c) {
    const v = this.wb.getValue(this.sheet, r, c);
    return v !== null && v !== undefined && v !== '';
  }

  drawBorder(r, c, border) {
    const { x, y, w, h } = this.cellRect(r, c);
    const widths = { thin: 1, hair: 1, dotted: 1, dashed: 1, medium: 2, mediumDashed: 2, thick: 3, double: 3 };
    const sides = {
      top: [x - 1, y - 0.5, x + w, y - 0.5],
      bottom: [x - 1, y + h - 0.5, x + w, y + h - 0.5],
      left: [x - 0.5, y - 1, x - 0.5, y + h],
      right: [x + w - 0.5, y - 1, x + w - 0.5, y + h],
    };
    for (const [side, kind] of Object.entries(border)) {
      if (!kind || !sides[side]) continue;
      const ctx = this.ctx;
      ctx.save();
      if (/dash|dot|hair/i.test(kind)) ctx.setLineDash(kind === 'dotted' || kind === 'hair' ? [1, 2] : [4, 2]);
      line(ctx, ...sides[side], '#000000', widths[kind] || 1);
      ctx.restore();
    }
  }

  drawSelection() {
    const ctx = this.ctx;
    const s = this.sel;
    const rect = this.rangeRect(s);

    if (s.r0 !== s.r1 || s.c0 !== s.c1) {
      ctx.fillStyle = 'rgba(33, 115, 70, 0.12)';
      ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
      const a = this.cellRect(this.active.r, this.active.c);
      // Repaint the active cell without the tint so it stands out, like Excel.
      ctx.save();
      ctx.beginPath();
      ctx.rect(a.x, a.y, a.w - 1, a.h - 1);
      ctx.clip();
      ctx.fillStyle = this.wb.getStyle(this.sheet, this.active.r, this.active.c)?.fill || '#ffffff';
      ctx.fillRect(a.x, a.y, a.w, a.h);
      this.drawText(this.active.r, this.active.c, this.active.c, this.active.c);
      ctx.restore();
    }

    if (this.edit?.refs) {
      this.edit.refs.forEach((ref, i) => {
        const rr = this.rangeRect(ref);
        const color = REF_COLORS[i % REF_COLORS.length];
        ctx.fillStyle = `${color}18`;
        ctx.fillRect(rr.x, rr.y, rr.w, rr.h);
        strokeRect(ctx, rr, color, 2);
      });
    }

    if (this.clip && this.clip.sheet === this.sheet) {
      ctx.save();
      ctx.setLineDash([5, 3]);
      strokeRect(ctx, this.rangeRect(this.clip.range), ACCENT, 2);
      ctx.restore();
    }

    if (this.fillTarget) {
      ctx.save();
      ctx.setLineDash([3, 3]);
      strokeRect(ctx, this.rangeRect(this.fillTarget), '#5f5f5f', 1);
      ctx.restore();
    }

    if (!this.edit) {
      strokeRect(ctx, rect, ACCENT, 2);
      // Fill handle
      ctx.fillStyle = ACCENT;
      ctx.fillRect(rect.x + rect.w - 4, rect.y + rect.h - 4, 7, 7);
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 1;
      ctx.strokeRect(rect.x + rect.w - 4.5, rect.y + rect.h - 4.5, 8, 8);
    } else {
      strokeRect(ctx, this.cellRect(this.active.r, this.active.c), ACCENT, 2);
    }
  }

  rangeRect(range) {
    const x0 = this.colX(range.c0);
    const y0 = this.rowY(range.r0);
    const x1 = this.colX(range.c1) + this.cols.size(range.c1);
    const y1 = this.rowY(range.r1) + this.rows.size(range.r1);
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  }

  drawHeaders(r0, r1, c0, c1) {
    const ctx = this.ctx;
    const hw = this.headerW;
    const s = this.sel;
    ctx.font = `12px 'Segoe UI', ${FONT_FALLBACK}`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';

    const drawCol = (c) => {
      const x = this.colX(c);
      const w = this.cols.size(c);
      if (!w) return;
      const selected = c >= s.c0 && c <= s.c1;
      const whole = selected && s.r0 === 0 && s.r1 >= this.rows.count - 1;
      ctx.fillStyle = whole ? '#c6e0d0' : selected ? '#d2d2d2' : '#f3f3f3';
      ctx.fillRect(x, 0, w, HEADER_H);
      ctx.fillStyle = selected ? ACCENT : '#444';
      ctx.fillText(colName(c), x + w / 2, HEADER_H / 2 + 1);
      line(ctx, x + w - 0.5, 0, x + w - 0.5, HEADER_H, '#c8c8c8', 1);
      if (selected) line(ctx, x, HEADER_H - 1, x + w, HEADER_H - 1, ACCENT, 2);
    };
    const drawRow = (r) => {
      const y = this.rowY(r);
      const h = this.rows.size(r);
      if (!h) return;
      const selected = r >= s.r0 && r <= s.r1;
      const whole = selected && s.c0 === 0 && s.c1 >= this.cols.count - 1;
      ctx.fillStyle = whole ? '#c6e0d0' : selected ? '#d2d2d2' : '#f3f3f3';
      ctx.fillRect(0, y, hw, h);
      ctx.fillStyle = selected ? ACCENT : '#444';
      ctx.fillText(String(r + 1), hw / 2, y + h / 2 + 1);
      line(ctx, 0, y + h - 0.5, hw, y + h - 0.5, '#c8c8c8', 1);
      if (selected) line(ctx, hw - 1, y, hw - 1, y + h, ACCENT, 2);
    };

    const fw = this.frozenW();
    const fh = this.frozenH();
    ctx.save();
    ctx.beginPath();
    ctx.rect(hw + fw, 0, this.width, HEADER_H);
    ctx.clip();
    for (let c = c0; c <= c1; c++) drawCol(c);
    ctx.restore();
    for (let c = 0; c < this.frozen.cols; c++) drawCol(c);

    ctx.save();
    ctx.beginPath();
    ctx.rect(0, HEADER_H + fh, hw, this.height);
    ctx.clip();
    for (let r = r0; r <= r1; r++) drawRow(r);
    ctx.restore();
    for (let r = 0; r < this.frozen.rows; r++) drawRow(r);

    line(ctx, 0, HEADER_H - 0.5, this.width, HEADER_H - 0.5, '#c8c8c8', 1);
    line(ctx, hw - 0.5, 0, hw - 0.5, this.height, '#c8c8c8', 1);
    // Corner: select all
    ctx.fillStyle = '#f3f3f3';
    ctx.fillRect(0, 0, hw - 1, HEADER_H - 1);
    ctx.fillStyle = '#b4b4b4';
    ctx.beginPath();
    ctx.moveTo(hw - 5, 5);
    ctx.lineTo(hw - 5, HEADER_H - 5);
    ctx.lineTo(hw - 5 - (HEADER_H - 10), HEADER_H - 5);
    ctx.fill();
    ctx.textAlign = 'start';
  }

  // ---------- Selection ----------

  select(r, c, { extend = false, keepAnchor = false } = {}) {
    this.tabStart = undefined;
    r = clamp(r, 0, MAX_ROWS - 1);
    c = clamp(c, 0, MAX_COLS - 1);
    this.rows.ensure(r + 50);
    this.cols.ensure(c + 10);
    if (extend) {
      this.sel = normRange({ r0: this.anchor.r, c0: this.anchor.c, r1: r, c1: c });
      this.cursor = { r, c };
    } else {
      this.active = { r, c };
      if (!keepAnchor) this.anchor = { r, c };
      this.sel = { r0: r, c0: c, r1: r, c1: c };
      this.cursor = { r, c };
    }
    this.updateSizer();
    this.scrollTo(this.cursor.r, this.cursor.c);
    this.draw();
    this.placeEditor();
    this.hooks.onSelection?.();
  }

  selectRange(range, active) {
    this.sel = normRange(range);
    this.active = active || { r: this.sel.r0, c: this.sel.c0 };
    this.anchor = { ...this.active };
    this.cursor = { r: this.sel.r1, c: this.sel.c1 };
    this.rows.ensure(this.sel.r1 + 1);
    this.cols.ensure(this.sel.c1 + 1);
    this.updateSizer();
    this.scrollTo(this.active.r, this.active.c);
    this.draw();
    this.placeEditor();
    this.hooks.onSelection?.();
  }

  selectAll() {
    this.selectRange({ r0: 0, c0: 0, r1: this.rows.count - 1, c1: this.cols.count - 1 }, this.active);
  }

  scrollTo(r, c) {
    const sc = this.scroller;
    if (r >= this.frozen.rows) {
      const top = this.rows.pos(r);
      const bottom = top + this.rows.size(r);
      const viewTop = sc.scrollTop + this.frozenH();
      const viewBottom = sc.scrollTop + this.height - HEADER_H;
      if (top < viewTop) sc.scrollTop = top - this.frozenH();
      else if (bottom > viewBottom) sc.scrollTop = bottom - (this.height - HEADER_H);
    }
    if (c >= this.frozen.cols) {
      const left = this.cols.pos(c);
      const right = left + this.cols.size(c);
      const viewLeft = sc.scrollLeft + this.frozenW();
      const viewRight = sc.scrollLeft + this.width - this.headerW;
      if (left < viewLeft) sc.scrollLeft = left - this.frozenW();
      else if (right > viewRight) sc.scrollLeft = right - (this.width - this.headerW);
    }
  }

  // Ctrl+arrow: jump to the edge of the current block of data.
  jump(r, c, dr, dc) {
    const filled = (rr, cc) => rr >= 0 && cc >= 0 && this.hasContent(rr, cc);
    const used = this.wb.usedSize(this.sheet);
    const limitR = Math.max(used.rows, r + 1);
    const limitC = Math.max(used.cols, c + 1);
    let nr = r + dr;
    let nc = c + dc;
    const inside = (rr, cc) => rr >= 0 && cc >= 0 && rr < limitR && cc < limitC;
    if (!inside(nr, nc)) return { r: clamp(nr, 0, dr > 0 ? this.rows.count - 1 : r), c: clamp(nc, 0, dc > 0 ? this.cols.count - 1 : c) };
    if (filled(r, c) && filled(nr, nc)) {
      while (inside(nr + dr, nc + dc) && filled(nr + dr, nc + dc)) {
        nr += dr;
        nc += dc;
      }
      return { r: nr, c: nc };
    }
    while (inside(nr, nc) && !filled(nr, nc)) {
      nr += dr;
      nc += dc;
    }
    if (!inside(nr, nc)) return { r: dr > 0 ? Math.max(r, limitR - 1) : dr < 0 ? 0 : r, c: dc > 0 ? Math.max(c, limitC - 1) : dc < 0 ? 0 : c };
    return { r: nr, c: nc };
  }

  move(dr, dc, { extend = false, jump = false } = {}) {
    const from = extend ? this.cursor || this.active : this.active;
    let target = { r: from.r + dr, c: from.c + dc };
    if (jump) target = this.jump(from.r, from.c, dr, dc);
    // Skip hidden rows and columns
    while (target.r > 0 && target.r < MAX_ROWS - 1 && !this.rows.size(target.r) && dr) target.r += dr;
    while (target.c > 0 && target.c < MAX_COLS - 1 && !this.cols.size(target.c) && dc) target.c += dc;
    this.select(target.r, target.c, { extend });
  }

  // Enter / Tab move inside a multi-cell selection without losing it.
  moveWithin(dr, dc) {
    const s = this.sel;
    if (s.r0 === s.r1 && s.c0 === s.c1) {
      // Typing a row with Tab, then Enter, goes back to the column you started in.
      if (dc) {
        const start = this.tabStart ?? this.active.c;
        this.move(0, dc);
        this.tabStart = start;
      } else if (this.tabStart !== undefined && dr > 0) {
        const c = this.tabStart;
        this.select(this.active.r + 1, c);
      } else this.move(dr, 0);
      return;
    }
    let { r, c } = this.active;
    if (dc) {
      c += dc;
      if (c > s.c1) [c, r] = [s.c0, r + 1 > s.r1 ? s.r0 : r + 1];
      if (c < s.c0) [c, r] = [s.c1, r - 1 < s.r0 ? s.r1 : r - 1];
    } else {
      r += dr;
      if (r > s.r1) [r, c] = [s.r0, c + 1 > s.c1 ? s.c0 : c + 1];
      if (r < s.r0) [r, c] = [s.r1, c - 1 < s.c0 ? s.c1 : c - 1];
    }
    this.active = { r, c };
    this.scrollTo(r, c);
    this.draw();
    this.placeEditor();
    this.hooks.onSelection?.();
  }

  // ---------- Editing ----------

  inputTextFor(r, c) {
    const wb = this.wb;
    if (wb.isFormula(this.sheet, r, c)) return wb.getInput(this.sheet, r, c);
    const value = wb.getValue(this.sheet, r, c);
    return editText(value, wb.getNumFmt(this.sheet, r, c)) ?? wb.getInput(this.sheet, r, c);
  }

  // `source` is where the user types: the in-cell editor or the formula bar.
  beginEdit(mode, text, source = this.input) {
    if (this.edit) return;
    const original = this.inputTextFor(this.active.r, this.active.c);
    this.edit = { mode, original, point: null, source };
    this.input.value = text ?? original;
    this.input.classList.add('editing');
    this.placeEditor();
    if (source === this.input) {
      if (mode === 'edit') this.input.setSelectionRange(this.input.value.length, this.input.value.length);
      this.input.focus();
    }
    this.updateRefs();
    this.draw();
    this.hooks.onEdit?.(this.input.value);
  }

  // Ends editing. Returns false if nothing was being edited.
  commitEdit() {
    if (!this.edit) return false;
    const text = this.input.value;
    const changed = text !== this.edit.original;
    this.endEdit();
    if (changed) {
      let input = text;
      if (input.startsWith('=')) input = closeParens(input);
      const { r, c } = this.active;
      this.wb.transact(() => {
        this.wb.enter(this.sheet, r, c, input);
        this.widenToFit(r, c);
      });
    }
    return true;
  }

  // Like Excel, a column still at its default width grows to show a typed number or date.
  widenToFit(r, c) {
    const id = this.sheet;
    if (this.wb.metaOf(id).colWidths.has(c) || this.wb.isFormula(id, r, c)) return;
    const value = this.wb.getValue(id, r, c);
    if (typeof value !== 'number') return;
    const fmt = this.wb.getNumFmt(id, r, c);
    if (fmt === 'General') return; // General numbers shrink their decimals instead
    this.ctx.font = this.fontFor(this.wb.getStyle(id, r, c));
    const needed = Math.ceil(this.ctx.measureText(formatValue(value, fmt)).width + 8);
    if (needed > this.cols.size(c)) this.wb.editMeta(id, 'colWidths', (m) => m.set(c, needed));
  }

  cancelEdit() {
    if (!this.edit) return;
    this.endEdit();
    this.hooks.onSelection?.();
  }

  endEdit() {
    this.edit = null;
    this.input.value = '';
    this.input.classList.remove('editing');
    this.placeEditor();
    this.draw();
    this.hooks.onEditEnd?.();
  }

  // Called when the formula bar changes the text.
  setEditText(text, source) {
    if (!this.edit) this.beginEdit('edit', text, source);
    this.input.value = text;
    this.edit.point = null;
    this.updateRefs();
    this.placeEditor();
    this.draw();
  }

  placeEditor() {
    const { r, c } = this.active;
    const rect = this.cellRect(r, c);
    const s = this.input.style;
    if (!this.edit) {
      // Parked over the active cell, invisible, so typing starts an edit there.
      s.left = `${clamp(rect.x, 0, this.width)}px`;
      s.top = `${clamp(rect.y, 0, this.height)}px`;
      s.width = '1px';
      s.height = '1px';
      return;
    }
    const style = this.wb.getStyle(this.sheet, r, c) || {};
    s.font = this.fontFor(style);
    s.color = style.color || '#000';
    s.background = style.fill || '#fff';
    s.textAlign = style.align && !this.input.value.startsWith('=') ? style.align : 'left';
    s.whiteSpace = style.wrap ? 'pre-wrap' : 'pre';
    this.ctx.font = this.fontFor(style);
    const longest = Math.max(...this.input.value.split('\n').map((t) => this.ctx.measureText(t).width));
    const maxW = Math.max(rect.w, this.width - rect.x - 4);
    const w = clamp(Math.max(rect.w, longest + 16), rect.w, maxW);
    const lines = this.input.value.split('\n').length;
    const lineH = (((style.size || 11) * 4) / 3) * 1.2;
    s.left = `${rect.x}px`;
    s.top = `${rect.y}px`;
    s.width = `${w}px`;
    s.height = `${Math.max(rect.h, lines * lineH + 4)}px`;
  }

  // Formula references to outline in colour while typing a formula.
  updateRefs() {
    if (!this.edit) return;
    const text = this.input.value;
    this.edit.refs = [];
    if (!text.startsWith('=')) return;
    const parts = text.split(/("(?:[^"]|"")*")/);
    for (let i = 0; i < parts.length; i += 2) {
      for (const m of parts[i].matchAll(/(?<![A-Za-z0-9_!.'])(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)(?![A-Za-z0-9_(!])/g)) {
        const ref = parseRef(m[1]);
        if (ref) this.edit.refs.push(ref);
      }
    }
  }

  // True when the caret sits where a cell reference could go (after "=", "(", "+", ",", …).
  canPoint() {
    if (!this.edit) return false;
    const src = this.edit.source;
    const text = src.value;
    if (!text.startsWith('=')) return false;
    const caret = src.selectionStart;
    if (this.edit.point && this.edit.point.end === caret) return true;
    if (src.selectionEnd !== caret) return false;
    return POINT_AFTER.test(text.slice(0, caret));
  }

  // Puts a reference to `range` into the formula at the caret (or replaces the last one pointed at).
  pointAt(range) {
    const src = this.edit.source;
    const text = src.value;
    const point = this.edit.point && this.edit.point.end === src.selectionStart ? this.edit.point : null;
    const start = point ? point.start : src.selectionStart;
    const end = point ? point.end : src.selectionEnd;
    const ref = rangeName(normRange(range));
    src.value = text.slice(0, start) + ref + text.slice(end);
    this.input.value = src.value;
    const caret = start + ref.length;
    src.setSelectionRange(caret, caret);
    this.edit.point = { start, end: caret, range: normRange(range), anchor: point?.anchor || { r: range.r0, c: range.c0 } };
    this.edit.mode = 'enter';
    this.updateRefs();
    this.placeEditor();
    this.draw();
    this.hooks.onEdit?.(this.input.value);
  }

  // ---------- Events ----------

  bindEvents() {
    this.scroller.addEventListener('scroll', () => {
      const sc = this.scroller;
      if (sc.scrollTop + this.height > sc.scrollHeight - 400) this.rows.ensure(this.rows.count + 500);
      if (sc.scrollLeft + this.width > sc.scrollWidth - 400) this.cols.ensure(this.cols.count + 20);
      this.updateSizer();
      this.draw();
      this.placeEditor();
    });
    this.canvas.addEventListener(
      'wheel',
      (e) => {
        e.preventDefault();
        const scale = e.deltaMode === 1 ? 20 : e.deltaMode === 2 ? this.height : 1;
        if (e.shiftKey && !e.deltaX) this.scroller.scrollLeft += e.deltaY * scale;
        else {
          this.scroller.scrollTop += e.deltaY * scale;
          this.scroller.scrollLeft += e.deltaX * scale;
        }
      },
      { passive: false },
    );
    this.canvas.addEventListener('mousedown', (e) => this.onMouseDown(e));
    this.canvas.addEventListener('mousemove', (e) => this.onHover(e));
    this.canvas.addEventListener('dblclick', (e) => this.onDoubleClick(e));
    this.canvas.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const hit = this.hitTest(e);
      if (hit.kind === 'cell' && !inRange(this.sel, hit.r, hit.c)) this.select(hit.r, hit.c);
      if (hit.kind === 'col' && !(hit.c >= this.sel.c0 && hit.c <= this.sel.c1 && this.sel.r0 === 0)) this.selectColumns(hit.c, hit.c);
      if (hit.kind === 'row' && !(hit.r >= this.sel.r0 && hit.r <= this.sel.r1 && this.sel.c0 === 0)) this.selectRows(hit.r, hit.r);
      this.hooks.onContextMenu?.(e, hit);
    });
    this.input.addEventListener('keydown', (e) => this.onKeyDown(e));
    this.input.addEventListener('input', () => {
      if (!this.edit) {
        const typed = this.input.value;
        this.input.value = '';
        this.beginEdit('enter', typed);
        return;
      }
      this.edit.point = null;
      this.updateRefs();
      this.placeEditor();
      this.draw();
      this.hooks.onEdit?.(this.input.value);
    });
    this.input.addEventListener('blur', () => {
      // Clicking the ribbon or formula bar keeps the edit; the page decides when it ends.
    });
  }

  hitTest(e) {
    const rect = this.canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const hw = this.headerW;
    if (x < hw && y < HEADER_H) return { kind: 'corner', x, y };
    if (y < HEADER_H) {
      const c = this.colAtX(x);
      const edgeX = this.colX(c) + this.cols.size(c);
      if (Math.abs(x - edgeX) <= 4) return { kind: 'col-edge', c, x, y };
      const left = this.colX(c);
      if (c > 0 && Math.abs(x - left) <= 4) return { kind: 'col-edge', c: c - 1, x, y };
      return { kind: 'col', c, x, y };
    }
    if (x < hw) {
      const r = this.rowAtY(y);
      const edgeY = this.rowY(r) + this.rows.size(r);
      if (Math.abs(y - edgeY) <= 3) return { kind: 'row-edge', r, x, y };
      const top = this.rowY(r);
      if (r > 0 && Math.abs(y - top) <= 3) return { kind: 'row-edge', r: r - 1, x, y };
      return { kind: 'row', r, x, y };
    }
    const r = this.rowAtY(y);
    const c = this.colAtX(x);
    const sr = this.rangeRect(this.sel);
    if (!this.edit && Math.abs(x - (sr.x + sr.w)) <= 5 && Math.abs(y - (sr.y + sr.h)) <= 5) return { kind: 'fill', r, c, x, y };
    return { kind: 'cell', r, c, x, y };
  }

  onHover(e) {
    if (this.dragging) return;
    const hit = this.hitTest(e);
    const cursors = { 'col-edge': 'col-resize', 'row-edge': 'row-resize', fill: 'crosshair', col: 's-resize', row: 'e-resize' };
    this.canvas.style.cursor = cursors[hit.kind] || 'cell';
  }

  selectColumns(a, b) {
    this.selectRange({ r0: 0, r1: this.rows.count - 1, c0: Math.min(a, b), c1: Math.max(a, b) }, { r: this.firstScrollRow(), c: a });
    this.active = { r: 0, c: a };
    this.sel = { r0: 0, r1: this.rows.count - 1, c0: Math.min(a, b), c1: Math.max(a, b) };
    this.draw();
  }

  selectRows(a, b) {
    this.selectRange({ r0: Math.min(a, b), r1: Math.max(a, b), c0: 0, c1: this.cols.count - 1 }, { r: a, c: this.firstScrollCol() });
    this.active = { r: a, c: 0 };
    this.sel = { r0: Math.min(a, b), r1: Math.max(a, b), c0: 0, c1: this.cols.count - 1 };
    this.draw();
  }

  onMouseDown(e) {
    if (e.button !== 0) return;
    const hit = this.hitTest(e);
    e.preventDefault();

    if (hit.kind === 'col-edge' || hit.kind === 'row-edge') return this.startResize(e, hit);

    // Pointing at cells while typing a formula
    if (this.edit && hit.kind === 'cell' && this.canPoint()) {
      const start = { r: hit.r, c: hit.c };
      this.pointAt({ r0: hit.r, c0: hit.c, r1: hit.r, c1: hit.c });
      this.drag(e, (ev) => {
        const h = this.hitTest(ev);
        if (h.kind !== 'cell') return;
        this.pointAt({ r0: start.r, c0: start.c, r1: h.r, c1: h.c });
      });
      this.edit.source.focus();
      return;
    }

    if (this.edit) this.commitEdit();
    this.input.focus();

    if (hit.kind === 'corner') return this.selectAll();
    if (hit.kind === 'fill') return this.startFill(e);
    if (hit.kind === 'col') {
      const from = e.shiftKey ? this.anchor.c : hit.c;
      this.selectColumns(from, hit.c);
      this.anchor = { r: 0, c: from };
      this.drag(e, (ev) => {
        const c = this.colAtX(ev.clientX - this.canvas.getBoundingClientRect().left);
        this.sel = { ...this.sel, c0: Math.min(from, c), c1: Math.max(from, c) };
        this.draw();
        this.hooks.onSelection?.();
      });
      return;
    }
    if (hit.kind === 'row') {
      const from = e.shiftKey ? this.anchor.r : hit.r;
      this.selectRows(from, hit.r);
      this.anchor = { r: from, c: 0 };
      this.drag(e, (ev) => {
        const r = this.rowAtY(ev.clientY - this.canvas.getBoundingClientRect().top);
        this.sel = { ...this.sel, r0: Math.min(from, r), r1: Math.max(from, r) };
        this.draw();
        this.hooks.onSelection?.();
      });
      return;
    }

    if (e.shiftKey) this.select(hit.r, hit.c, { extend: true });
    else this.select(hit.r, hit.c);
    this.drag(e, (ev) => {
      const h = this.hitTest(ev);
      if (h.kind === 'cell' || h.kind === 'fill') this.select(h.r, h.c, { extend: true });
    });
  }

  // Tracks the mouse until release, auto-scrolling near the edges.
  drag(e, onMove, onEnd) {
    this.dragging = true;
    let last = e;
    const timer = setInterval(() => {
      const rect = this.canvas.getBoundingClientRect();
      const sc = this.scroller;
      let moved = false;
      if (last.clientY > rect.bottom - 4) (sc.scrollTop += 20), (moved = true);
      else if (last.clientY < rect.top + HEADER_H && sc.scrollTop > 0) (sc.scrollTop -= 20), (moved = true);
      if (last.clientX > rect.right - 4) (sc.scrollLeft += 30), (moved = true);
      else if (last.clientX < rect.left + this.headerW && sc.scrollLeft > 0) (sc.scrollLeft -= 30), (moved = true);
      if (moved) onMove(last);
    }, 50);
    const move = (ev) => {
      last = ev;
      onMove(ev);
    };
    const up = (ev) => {
      clearInterval(timer);
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      this.dragging = false;
      onEnd?.(ev);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }

  startResize(e, hit) {
    const isCol = hit.kind === 'col-edge';
    const index = isCol ? hit.c : hit.r;
    const startPos = isCol ? e.clientX : e.clientY;
    const startSize = isCol ? this.cols.size(index) : this.rows.size(index);
    const targets = isCol
      ? this.sel.r0 === 0 && this.sel.r1 >= this.rows.count - 1 && index >= this.sel.c0 && index <= this.sel.c1
        ? range(this.sel.c0, this.sel.c1)
        : [index]
      : this.sel.c0 === 0 && this.sel.c1 >= this.cols.count - 1 && index >= this.sel.r0 && index <= this.sel.r1
        ? range(this.sel.r0, this.sel.r1)
        : [index];
    let size = startSize;
    this.resizeGuide = null;
    this.drag(
      e,
      (ev) => {
        size = Math.max(0, startSize + (isCol ? ev.clientX : ev.clientY) - startPos);
        this.draw();
        const ctx = this.ctx;
        if (isCol) {
          const x = this.colX(index) + size;
          line(ctx, x, 0, x, this.height, ACCENT, 1);
        } else {
          const y = this.rowY(index) + size;
          line(ctx, 0, y, this.width, y, ACCENT, 1);
        }
      },
      () => {
        if (size === startSize) return this.draw();
        if (isCol) this.wb.setColWidth(this.sheet, targets, size);
        else this.wb.setRowHeight(this.sheet, targets, size);
      },
    );
  }

  startFill(e) {
    const src = { ...this.sel };
    this.drag(
      e,
      (ev) => {
        const h = this.hitTest(ev);
        if (h.kind !== 'cell' && h.kind !== 'fill') return;
        const downDist = h.r > src.r1 ? h.r - src.r1 : h.r < src.r0 ? src.r0 - h.r : 0;
        const rightDist = h.c > src.c1 ? h.c - src.c1 : h.c < src.c0 ? src.c0 - h.c : 0;
        if (!downDist && !rightDist) this.fillTarget = null;
        else if (downDist >= rightDist) this.fillTarget = { ...src, r0: Math.min(src.r0, h.r), r1: Math.max(src.r1, h.r) };
        else this.fillTarget = { ...src, c0: Math.min(src.c0, h.c), c1: Math.max(src.c1, h.c) };
        this.draw();
      },
      () => {
        const target = this.fillTarget;
        this.fillTarget = null;
        if (!target) return this.draw();
        this.wb.fill(this.sheet, src, target);
        this.selectRange(target, this.active);
      },
    );
  }

  onDoubleClick(e) {
    const hit = this.hitTest(e);
    if (hit.kind === 'col-edge') return this.autofitColumn(hit.c);
    if (hit.kind === 'row-edge') return this.wb.setRowHeight(this.sheet, [hit.r], 20);
    if (hit.kind === 'cell') {
      this.select(hit.r, hit.c);
      this.beginEdit('edit');
    }
  }

  autofitColumn(c) {
    const targets = this.sel.r0 === 0 && this.sel.r1 >= this.rows.count - 1 && c >= this.sel.c0 && c <= this.sel.c1 ? range(this.sel.c0, this.sel.c1) : [c];
    const used = this.wb.usedSize(this.sheet);
    const widths = targets.map((col) => {
      let max = 0;
      for (let r = 0; r < used.rows; r++) {
        const value = this.wb.getValue(this.sheet, r, col);
        if (value === null || value === undefined || value === '') continue;
        const style = this.wb.getStyle(this.sheet, r, col);
        this.ctx.font = this.fontFor(style);
        max = Math.max(max, this.ctx.measureText(formatValue(value, this.wb.getNumFmt(this.sheet, r, col))).width);
      }
      return max ? Math.ceil(max + 10) : 64;
    });
    this.wb.transact(() =>
      this.wb.editMeta(this.sheet, 'colWidths', (m) => {
        targets.forEach((col, i) => (widths[i] === 64 ? m.delete(col) : m.set(col, widths[i])));
      }),
    );
  }

  onKeyDown(e) {
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;

    if (this.edit) {
      if (k === 'Enter' && e.altKey) {
        e.preventDefault();
        const { selectionStart: a, selectionEnd: b, value } = this.input;
        this.input.value = `${value.slice(0, a)}\n${value.slice(b)}`;
        this.input.setSelectionRange(a + 1, a + 1);
        this.input.dispatchEvent(new Event('input'));
        return;
      }
      if (k === 'Enter' || k === 'Tab') {
        e.preventDefault();
        if (this.hooks.onAutocompleteAccept?.(k)) return;
        this.commitEdit();
        if (k === 'Enter' && mod) return;
        if (k === 'Tab') this.moveWithin(0, e.shiftKey ? -1 : 1);
        else this.moveWithin(e.shiftKey ? -1 : 1, 0);
        return;
      }
      if (k === 'Escape') {
        e.preventDefault();
        if (this.hooks.onAutocompleteCancel?.()) return;
        this.cancelEdit();
        return;
      }
      if (this.hooks.onAutocompleteKey?.(e)) return;
      if (k === 'F2') {
        e.preventDefault();
        this.edit.mode = this.edit.mode === 'edit' ? 'enter' : 'edit';
        return;
      }
      const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (arrows[k] && this.edit.mode === 'enter') {
        e.preventDefault();
        const [dr, dc] = arrows[k];
        if (this.canPoint()) {
          // Arrow keys move the reference being pointed at, as in Excel.
          const p = this.edit.point;
          const base = p ? (e.shiftKey ? { r: p.range.r1, c: p.range.c1 } : p.anchor) : this.active;
          const target = { r: clamp(base.r + dr, 0, MAX_ROWS - 1), c: clamp(base.c + dc, 0, MAX_COLS - 1) };
          const anchor = p && e.shiftKey ? p.anchor : target;
          this.pointAt({ r0: anchor.r, c0: anchor.c, r1: target.r, c1: target.c });
          if (this.edit.point) this.edit.point.anchor = anchor;
          this.scrollTo(target.r, target.c);
          return;
        }
        this.commitEdit();
        this.move(dr, dc);
      }
      return;
    }

    // Not editing
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) {
      e.preventDefault();
      this.move(...arrows[k], { extend: e.shiftKey, jump: mod });
      return;
    }
    switch (k) {
      case 'Enter':
        e.preventDefault();
        this.moveWithin(e.shiftKey ? -1 : 1, 0);
        return;
      case 'Tab':
        e.preventDefault();
        this.moveWithin(0, e.shiftKey ? -1 : 1);
        return;
      case 'F2':
        e.preventDefault();
        this.beginEdit('edit');
        return;
      case 'Delete':
        e.preventDefault();
        this.wb.clear(this.sheet, this.sel, 'contents');
        return;
      case 'Backspace':
        e.preventDefault();
        this.beginEdit('enter', '');
        return;
      case 'Escape':
        if (this.clip) {
          this.clip = null;
          this.hooks.onClipCleared?.();
          this.draw();
        }
        return;
      case 'Home':
        e.preventDefault();
        if (mod) this.select(this.frozen.rows, this.frozen.cols, { extend: e.shiftKey });
        else this.select(this.active.r, 0, { extend: e.shiftKey });
        return;
      case 'End':
        if (mod) {
          e.preventDefault();
          const used = this.wb.usedSize(this.sheet);
          this.select(Math.max(0, used.rows - 1), Math.max(0, used.cols - 1), { extend: e.shiftKey });
        }
        return;
      case 'PageDown':
      case 'PageUp': {
        e.preventDefault();
        if (mod) return this.hooks.onSheetStep?.(k === 'PageDown' ? 1 : -1);
        const visibleRows = Math.max(1, this.lastVisibleRow() - this.firstScrollRow());
        const dir = k === 'PageDown' ? 1 : -1;
        this.scroller.scrollTop += dir * (this.height - HEADER_H - this.frozenH());
        this.move(dir * visibleRows, 0, { extend: e.shiftKey });
        return;
      }
      default:
    }
    if (e.key === ' ' && mod && !e.shiftKey) {
      e.preventDefault();
      this.selectColumns(this.sel.c0, this.sel.c1);
      return;
    }
    if (e.key === ' ' && e.shiftKey && !mod) {
      e.preventDefault();
      this.selectRows(this.sel.r0, this.sel.r1);
      return;
    }
    if (mod) this.hooks.onShortcut?.(e);
  }

  // The selection as a rectangle clipped to cells in use, for scanning.
  usedSelection() {
    const used = this.wb.usedSize(this.sheet);
    return {
      r0: this.sel.r0,
      c0: this.sel.c0,
      r1: Math.min(this.sel.r1, Math.max(this.sel.r0, used.rows - 1)),
      c1: Math.min(this.sel.c1, Math.max(this.sel.c0, used.cols - 1)),
    };
  }

  activeName() {
    return cellName(this.active.r, this.active.c);
  }
}

// ---------- Helpers ----------

function el(tag, className) {
  const node = document.createElement(tag);
  node.className = className;
  return node;
}

function clamp(n, lo, hi) {
  return Math.min(hi, Math.max(lo, n));
}

function range(a, b) {
  return Array.from({ length: b - a + 1 }, (_, i) => a + i);
}

function inRange(s, r, c) {
  return r >= s.r0 && r <= s.r1 && c >= s.c0 && c <= s.c1;
}

function line(ctx, x0, y0, x1, y1, color, width) {
  ctx.beginPath();
  ctx.strokeStyle = color;
  ctx.lineWidth = width;
  ctx.moveTo(x0, y0);
  ctx.lineTo(x1, y1);
  ctx.stroke();
}

function strokeRect(ctx, rect, color, width) {
  ctx.strokeStyle = color;
  ctx.lineWidth = width;
  const inset = width / 2;
  ctx.strokeRect(rect.x - 1 + inset, rect.y - 1 + inset, rect.w + 1 - width + 1, rect.h + 1 - width + 1);
}

function wrapLines(text, width, measure) {
  const out = [];
  for (const paragraph of text.split('\n')) {
    let current = '';
    for (const word of paragraph.split(/(\s+)/)) {
      const next = current + word;
      if (current && measure(next.trimEnd()) > width) {
        out.push(current.trimEnd());
        current = word.trimStart();
      } else current = next;
    }
    out.push(current);
  }
  return out;
}

// Excel adds missing closing brackets when you press Enter.
function closeParens(formula) {
  let depth = 0;
  let inString = false;
  for (const ch of formula) {
    if (ch === '"') inString = !inString;
    else if (!inString && ch === '(') depth++;
    else if (!inString && ch === ')') depth--;
  }
  return depth > 0 && !inString ? formula + ')'.repeat(depth) : formula;
}
