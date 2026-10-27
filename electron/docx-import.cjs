// Reads a .docx file into the editor's document format (TipTap/ProseMirror JSON).
// Covers what the editor can show: paragraph styles, alignment, lists, character
// formatting, links, tables, pictures and page breaks. Anything else (headers,
// footers, text boxes, comments, tracked-change markup) is skipped.
const JSZip = require('jszip');
const { DOMParser } = require('@xmldom/xmldom');

const EMU_PER_PX = 9525;
const IMAGE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };
const ALIGN = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' };
const HIGHLIGHT_NONE = new Set(['none', '']);

async function importDocx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await readText(zip, 'word/document.xml');
  if (!documentXml) throw new Error('This file is not a Word document (word/document.xml is missing).');

  const ctx = {
    zip,
    rels: parseRels(await readText(zip, 'word/_rels/document.xml.rels')),
    styles: parseStyles(await readText(zip, 'word/styles.xml')),
    numbering: parseNumbering(await readText(zip, 'word/numbering.xml')),
    images: new Map(),
  };

  const body = child(parseXml(documentXml).documentElement, 'body');
  const content = await convertBody(body ? kids(body) : [], ctx);
  return { type: 'doc', content: content.length ? content : [{ type: 'paragraph' }] };
}

// ---------- XML helpers (match on local names so namespace prefixes don't matter) ----------
function parseXml(text) {
  return new DOMParser().parseFromString(text, 'text/xml');
}

function kids(el) {
  const out = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n);
  return out;
}

function local(el) {
  return el.localName || el.nodeName.replace(/^.*:/, '');
}

function child(el, name) {
  return el && kids(el).find((k) => local(k) === name);
}

function descendants(el, name) {
  const out = [];
  const walk = (node) => {
    for (const k of kids(node)) {
      if (local(k) === name) out.push(k);
      walk(k);
    }
  };
  if (el) walk(el);
  return out;
}

function attr(el, name) {
  if (!el) return null;
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if ((a.localName || a.name.replace(/^.*:/, '')) === name) return a.value;
  }
  return null;
}

function val(el) {
  return attr(el, 'val');
}

// <w:b/> is on, <w:b w:val="false"/> or "0" is off.
function isOn(el) {
  if (!el) return false;
  const v = val(el);
  return v == null || !['false', '0', 'off', 'none'].includes(v);
}

async function readText(zip, path) {
  const file = zip.file(path);
  return file ? file.async('string') : null;
}

// ---------- Package parts ----------
function parseRels(xml) {
  const rels = new Map();
  if (!xml) return rels;
  for (const r of descendants(parseXml(xml).documentElement, 'Relationship')) {
    rels.set(attr(r, 'Id'), { target: attr(r, 'Target'), external: attr(r, 'TargetMode') === 'External' });
  }
  return rels;
}

// styleId -> { name, numId, ilvl } (style ids are localised in some Word versions; names are not)
function parseStyles(xml) {
  const styles = new Map();
  if (!xml) return styles;
  for (const s of descendants(parseXml(xml).documentElement, 'style')) {
    const pPr = child(s, 'pPr');
    const numPr = child(pPr, 'numPr');
    styles.set(attr(s, 'styleId'), {
      name: (val(child(s, 'name')) || '').toLowerCase(),
      numId: numPr ? val(child(numPr, 'numId')) : null,
      ilvl: numPr ? Number(val(child(numPr, 'ilvl')) || 0) : 0,
    });
  }
  return styles;
}

// numId -> (level -> 'bullet' | 'ordered')
function parseNumbering(xml) {
  const formats = new Map();
  if (!xml) return formats;
  const root = parseXml(xml).documentElement;
  const abstract = new Map();
  for (const a of descendants(root, 'abstractNum')) {
    const levels = new Map();
    for (const lvl of descendants(a, 'lvl')) {
      levels.set(Number(attr(lvl, 'ilvl')), val(child(lvl, 'numFmt')) === 'bullet' ? 'bullet' : 'ordered');
    }
    abstract.set(attr(a, 'abstractNumId'), levels);
  }
  for (const n of kids(root).filter((k) => local(k) === 'num')) {
    formats.set(attr(n, 'numId'), abstract.get(val(child(n, 'abstractNumId'))) || new Map());
  }
  return formats;
}

// ---------- Body ----------
async function convertBody(elements, ctx) {
  const blocks = [];
  const lists = new ListBuilder(blocks);

  for (const el of elements) {
    switch (local(el)) {
      case 'p': {
        const { before, block, list, extras } = await convertParagraph(el, ctx);
        if (before.length) {
          lists.reset();
          blocks.push(...before);
        }
        if (block) {
          if (list) lists.add(block, list);
          else {
            lists.reset();
            blocks.push(block);
          }
        }
        if (extras.length) {
          lists.reset();
          blocks.push(...extras);
        }
        break;
      }
      case 'tbl':
        lists.reset();
        blocks.push(await convertTable(el, ctx));
        break;
      case 'sdt': {
        // Content controls wrap ordinary paragraphs and tables.
        const inner = child(el, 'sdtContent');
        if (inner) {
          lists.reset();
          blocks.push(...(await convertBody(kids(inner), ctx)));
        }
        break;
      }
      default:
        break;
    }
  }
  return blocks;
}

// Word stores lists as flat numbered paragraphs; the editor nests them.
class ListBuilder {
  constructor(blocks) {
    this.blocks = blocks;
    this.stack = [];
  }

  reset() {
    this.stack = [];
  }

  add(paragraph, { type, level, numId }) {
    while (this.stack.length && this.top.level > level) this.stack.pop();

    const top = this.top;
    if (top && top.level === level && (top.type !== type || (level === 0 && top.numId !== numId))) {
      this.stack.pop();
    }

    if (!this.stack.length) {
      this.push(type, level, numId, this.blocks);
    } else if (this.top.level < level) {
      const items = this.top.list.content;
      if (!items.length) items.push({ type: 'listItem', content: [{ type: 'paragraph' }] });
      this.push(type, level, numId, items[items.length - 1].content);
    }
    this.top.list.content.push({ type: 'listItem', content: [paragraph] });
  }

  push(type, level, numId, container) {
    const list = { type: type === 'bullet' ? 'bulletList' : 'orderedList', content: [] };
    container.push(list);
    this.stack.push({ list, type, level, numId });
  }

  get top() {
    return this.stack[this.stack.length - 1];
  }
}

async function convertParagraph(p, ctx) {
  const pPr = child(p, 'pPr');
  const styleId = val(child(pPr, 'pStyle'));
  const style = ctx.styles.get(styleId) || { name: (styleId || '').toLowerCase() };
  const align = ALIGN[val(child(pPr, 'jc'))];

  const runs = { inlines: [], extras: [], pageBreakBefore: false, pageBreakAfter: false };
  await convertInlineChildren(kids(p), ctx, [], runs);

  let block;
  const headingMatch = style.name.match(/^heading\s*(\d)$/);
  if (headingMatch && Number(headingMatch[1]) <= 6) {
    block = { type: 'heading', attrs: { level: Number(headingMatch[1]) } };
  } else if (style.name === 'title') {
    block = { type: 'heading', attrs: { level: 1 } };
  } else {
    block = { type: 'paragraph' };
  }
  if (align) block.attrs = { ...block.attrs, textAlign: align };
  if (runs.inlines.length) block.content = trimHardBreaks(runs.inlines);

  // A paragraph that only held a picture or a page break becomes just that.
  const onlyExtras = !runs.inlines.length && (runs.extras.length || runs.pageBreakBefore || runs.pageBreakAfter);
  const before = runs.pageBreakBefore ? [{ type: 'pageBreak' }] : [];
  const extras = [...runs.extras];
  if (runs.pageBreakAfter) extras.push({ type: 'pageBreak' });

  let list = null;
  if (block.type === 'paragraph' && !onlyExtras) list = listInfo(pPr, style, ctx);
  return { before, block: onlyExtras ? null : block, list, extras };
}

function listInfo(pPr, style, ctx) {
  const numPr = child(pPr, 'numPr');
  let numId = numPr ? val(child(numPr, 'numId')) : style.numId;
  let level = numPr ? Number(val(child(numPr, 'ilvl')) || 0) : style.ilvl || 0;
  if (numId === '0') return null; // numbering explicitly removed

  if (numId) {
    const formats = ctx.numbering.get(numId);
    if (!formats) return null;
    return { type: formats.get(level) || 'bullet', level, numId };
  }
  // Older documents use the built-in list styles without numbering properties.
  const m = style.name.match(/^list (bullet|number)(?:\s*(\d))?$/);
  if (m) return { type: m[1] === 'bullet' ? 'bullet' : 'ordered', level: m[2] ? Number(m[2]) - 1 : 0, numId: style.name };
  return null;
}

function trimHardBreaks(inlines) {
  while (inlines.length && inlines[inlines.length - 1].type === 'hardBreak') inlines.pop();
  return inlines;
}

async function convertInlineChildren(elements, ctx, marks, out) {
  for (const el of elements) {
    switch (local(el)) {
      case 'r':
        await convertRun(el, ctx, marks, out);
        break;
      case 'hyperlink': {
        const rel = ctx.rels.get(attr(el, 'id'));
        const href = rel && rel.external ? rel.target : null;
        const linkMarks = href ? [...marks, { type: 'link', attrs: { href } }] : marks;
        await convertInlineChildren(kids(el), ctx, linkMarks, out);
        break;
      }
      case 'ins':
      case 'smartTag':
      case 'fldSimple':
      case 'customXml':
        await convertInlineChildren(kids(el), ctx, marks, out);
        break;
      case 'sdt': {
        const inner = child(el, 'sdtContent');
        if (inner) await convertInlineChildren(kids(inner), ctx, marks, out);
        break;
      }
      default:
        break; // pPr, bookmarks, deleted text, comments
    }
  }
}

async function convertRun(r, ctx, inherited, out) {
  const marks = [...inherited, ...runMarks(child(r, 'rPr'))];
  let text = '';
  const flush = () => {
    if (text) out.inlines.push(marks.length ? { type: 'text', text, marks } : { type: 'text', text });
    text = '';
  };

  for (const el of kids(r)) {
    switch (local(el)) {
      case 't':
        text += el.textContent;
        break;
      case 'tab':
        text += '\t';
        break;
      case 'noBreakHyphen':
        text += '‑';
        break;
      case 'br':
        flush();
        if (attr(el, 'type') !== 'page') out.inlines.push({ type: 'hardBreak' });
        // Word often puts the break at the start of the next page's first paragraph.
        else if (out.inlines.length || out.extras.length) out.pageBreakAfter = true;
        else out.pageBreakBefore = true;
        break;
      case 'drawing': {
        flush();
        const image = await convertDrawing(el, ctx);
        if (image) out.extras.push(image);
        break;
      }
      default:
        break;
    }
  }
  flush();
}

function runMarks(rPr) {
  if (!rPr) return [];
  const marks = [];
  if (isOn(child(rPr, 'b'))) marks.push({ type: 'bold' });
  if (isOn(child(rPr, 'i'))) marks.push({ type: 'italic' });
  const u = child(rPr, 'u');
  if (u && val(u) !== 'none') marks.push({ type: 'underline' });
  if (isOn(child(rPr, 'strike')) || isOn(child(rPr, 'dstrike'))) marks.push({ type: 'strike' });
  const highlight = val(child(rPr, 'highlight'));
  if (highlight && !HIGHLIGHT_NONE.has(highlight)) marks.push({ type: 'highlight' });

  const style = {};
  const font = attr(child(rPr, 'rFonts'), 'ascii') || attr(child(rPr, 'rFonts'), 'hAnsi');
  if (font) style.fontFamily = font;
  const size = Number(val(child(rPr, 'sz')));
  if (size) style.fontSize = `${size / 2}pt`;
  const color = val(child(rPr, 'color'));
  if (color && /^[0-9a-f]{6}$/i.test(color)) style.color = `#${color.toLowerCase()}`;
  if (Object.keys(style).length) {
    marks.push({ type: 'textStyle', attrs: { fontFamily: null, fontSize: null, color: null, ...style } });
  }
  return marks;
}

async function convertDrawing(drawing, ctx) {
  const blip = descendants(drawing, 'blip')[0];
  const rel = blip && ctx.rels.get(attr(blip, 'embed'));
  if (!rel || rel.external) return null;

  const path = resolvePart('word/', rel.target);
  const ext = path.split('.').pop().toLowerCase();
  const mime = IMAGE_TYPES[ext];
  const file = mime && ctx.zip.file(path);
  if (!file) return null;

  if (!ctx.images.has(path)) ctx.images.set(path, `data:${mime};base64,${await file.async('base64')}`);

  const extent = descendants(drawing, 'extent')[0];
  const width = Math.round(Number(attr(extent, 'cx')) / EMU_PER_PX) || null;
  const height = Math.round(Number(attr(extent, 'cy')) / EMU_PER_PX) || null;
  const docPr = descendants(drawing, 'docPr')[0];
  return {
    type: 'image',
    attrs: { src: ctx.images.get(path), alt: attr(docPr, 'descr') || null, title: null, width, height },
  };
}

function resolvePart(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (base + target).split('/');
  const out = [];
  for (const part of parts) {
    if (part === '..') out.pop();
    else if (part && part !== '.') out.push(part);
  }
  return out.join('/');
}

// ---------- Tables ----------
async function convertTable(tbl, ctx) {
  const rows = [];
  // Cells waiting for vertical-merge continuations, by starting grid column.
  const openMerges = new Map();

  for (const tr of kids(tbl).filter((k) => local(k) === 'tr')) {
    const header = isOn(child(child(tr, 'trPr'), 'tblHeader'));
    const cells = [];
    let column = Number(val(child(child(tr, 'trPr'), 'gridBefore')) || 0);

    for (const tc of kids(tr).filter((k) => local(k) === 'tc')) {
      const tcPr = child(tc, 'tcPr');
      const colspan = Number(val(child(tcPr, 'gridSpan')) || 1);
      const vMerge = child(tcPr, 'vMerge');
      const vMergeVal = vMerge ? val(vMerge) || 'continue' : null;

      if (vMergeVal === 'continue' && openMerges.has(column)) {
        openMerges.get(column).attrs.rowspan += 1;
        column += colspan;
        continue;
      }

      let content = await convertBody(kids(tc), ctx);
      if (!content.length) content = [{ type: 'paragraph' }];
      const cell = {
        type: header ? 'tableHeader' : 'tableCell',
        attrs: { colspan, rowspan: 1, colwidth: null },
        content,
      };
      if (vMergeVal === 'restart') openMerges.set(column, cell);
      else openMerges.delete(column);
      cells.push(cell);
      column += colspan;
    }
    if (cells.length) rows.push({ type: 'tableRow', content: cells });
  }
  return rows.length ? { type: 'table', content: rows } : { type: 'paragraph' };
}

module.exports = { importDocx };
