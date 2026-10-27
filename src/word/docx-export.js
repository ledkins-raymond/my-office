// Converts the editor's document (TipTap/ProseMirror JSON) into a .docx file.
// Kept free of browser APIs so the tests can run it in Node.
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  LineRuleType,
  PageBreak,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';

// US Letter with 1-inch margins, in twips (1/1440 inch).
const PAGE = { width: 12240, height: 15840, margin: 1440 };
const CONTENT_WIDTH_TWIPS = PAGE.width - 2 * PAGE.margin;
const CONTENT_WIDTH_PX = Math.round((CONTENT_WIDTH_TWIPS / 1440) * 96);

const HEADINGS = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
};

const ALIGNMENTS = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};

const ORDERED_FORMATS = [
  LevelFormat.DECIMAL,
  LevelFormat.LOWER_LETTER,
  LevelFormat.LOWER_ROMAN,
];

/**
 * @param {object} doc TipTap JSON (editor.getJSON())
 * @param {{ getImageSize?: (src: string) => Promise<{width: number, height: number}> }} options
 * @returns {Promise<Document>}
 */
export async function buildDocx(doc, { getImageSize } = {}) {
  const ctx = { getImageSize, orderedInstance: 0 };
  const children = await convertBlocks(doc.content || [], ctx, {});

  return new Document({
    creator: 'My Office',
    styles: {
      default: {
        document: {
          run: { font: 'Calibri', size: 22 },
          paragraph: { spacing: { after: 160, line: 259, lineRule: LineRuleType.AUTO } },
        },
        heading1: {
          run: { font: 'Calibri Light', size: 32, color: '2F5496' },
          paragraph: { spacing: { before: 240, after: 0 } },
        },
        heading2: {
          run: { font: 'Calibri Light', size: 26, color: '2F5496' },
          paragraph: { spacing: { before: 40, after: 0 } },
        },
        heading3: {
          run: { font: 'Calibri Light', size: 24, color: '1F3763' },
          paragraph: { spacing: { before: 40, after: 0 } },
        },
      },
    },
    numbering: {
      config: [
        {
          reference: 'ordered',
          levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            format: ORDERED_FORMATS[level % ORDERED_FORMATS.length],
            text: `%${level + 1}.`,
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
          })),
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE.width, height: PAGE.height },
            margin: { top: PAGE.margin, right: PAGE.margin, bottom: PAGE.margin, left: PAGE.margin },
          },
        },
        children: children.length ? children : [new Paragraph({})],
      },
    ],
  });
}

// `para` carries paragraph options inherited from the surrounding block,
// such as list membership or a blockquote indent.
async function convertBlocks(nodes, ctx, para) {
  const out = [];
  for (const node of nodes) {
    out.push(...(await convertBlock(node, ctx, para)));
  }
  return out;
}

async function convertBlock(node, ctx, para) {
  const attrs = node.attrs || {};
  switch (node.type) {
    case 'paragraph':
      return [
        new Paragraph({
          ...para,
          alignment: ALIGNMENTS[attrs.textAlign],
          children: await convertInlines(node.content || [], ctx),
        }),
      ];

    case 'heading':
      return [
        new Paragraph({
          ...para,
          heading: HEADINGS[attrs.level] || HeadingLevel.HEADING_1,
          alignment: ALIGNMENTS[attrs.textAlign],
          children: await convertInlines(node.content || [], ctx),
        }),
      ];

    case 'bulletList':
    case 'orderedList':
      return convertList(node, ctx, para);

    case 'blockquote': {
      const left = (para.indent?.left || 0) + 720;
      return convertBlocks(node.content || [], ctx, { ...para, indent: { left } });
    }

    case 'codeBlock': {
      const text = (node.content || []).map((t) => t.text || '').join('');
      return text.split('\n').map(
        (line) =>
          new Paragraph({
            ...para,
            spacing: { after: 0 },
            children: [new TextRun({ text: line, font: 'Courier New', size: 20 })],
          }),
      );
    }

    case 'horizontalRule':
      return [
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'A0A0A0', space: 1 } },
        }),
      ];

    case 'pageBreak':
      return [new Paragraph({ children: [new PageBreak()] })];

    case 'image': {
      const run = await convertImage(attrs, ctx);
      return [
        new Paragraph({
          ...para,
          alignment: ALIGNMENTS[attrs.textAlign],
          children: run ? [run] : [new TextRun({ text: '[image]', italics: true })],
        }),
      ];
    }

    case 'table':
      return [await convertTable(node, ctx)];

    default:
      // Unknown block: keep its text rather than dropping it.
      if (node.content) return convertBlocks(node.content, ctx, para);
      return [];
  }
}

async function convertList(list, ctx, para) {
  const level = para.listLevel ?? 0;
  const ordered = list.type === 'orderedList';
  const instance = ordered ? ++ctx.orderedInstance : undefined;
  const out = [];

  for (const item of list.content || []) {
    let first = true;
    for (const child of item.content || []) {
      if (child.type === 'bulletList' || child.type === 'orderedList') {
        out.push(...(await convertList(child, ctx, { listLevel: level + 1 })));
        continue;
      }
      let childPara;
      if (first) {
        // Contextual spacing drops the gap between items, as Word's List Paragraph style does.
        childPara = ordered
          ? { numbering: { reference: 'ordered', level, instance }, contextualSpacing: true }
          : { bullet: { level }, contextualSpacing: true };
      } else {
        // Continuation paragraphs line up with the item text.
        childPara = { indent: { left: 720 * (level + 1) }, contextualSpacing: true };
      }
      out.push(...(await convertBlock(child, ctx, childPara)));
      first = false;
    }
  }
  return out;
}

async function convertTable(node, ctx) {
  const rows = node.content || [];
  const columnCount = Math.max(
    1,
    ...rows.map((r) => (r.content || []).reduce((n, c) => n + (c.attrs?.colspan || 1), 0)),
  );
  const columnWidth = Math.floor(CONTENT_WIDTH_TWIPS / columnCount);

  const tableRows = [];
  for (const row of rows) {
    const cells = [];
    const isHeaderRow = (row.content || []).every((c) => c.type === 'tableHeader');
    for (const cell of row.content || []) {
      const colspan = cell.attrs?.colspan || 1;
      let children = await convertBlocks(cell.content || [], ctx, {});
      if (!children.length) children = [new Paragraph({})];
      cells.push(
        new TableCell({
          children,
          columnSpan: colspan > 1 ? colspan : undefined,
          rowSpan: (cell.attrs?.rowspan || 1) > 1 ? cell.attrs.rowspan : undefined,
          width: { size: columnWidth * colspan, type: WidthType.DXA },
          shading:
            cell.type === 'tableHeader'
              ? { type: ShadingType.CLEAR, color: 'auto', fill: 'F2F2F2' }
              : undefined,
        }),
      );
    }
    tableRows.push(new TableRow({ children: cells, tableHeader: isHeaderRow || undefined }));
  }

  return new Table({
    rows: tableRows,
    width: { size: CONTENT_WIDTH_TWIPS, type: WidthType.DXA },
    columnWidths: Array(columnCount).fill(columnWidth),
  });
}

async function convertInlines(nodes, ctx) {
  const out = [];
  for (const node of nodes) {
    if (node.type === 'text') {
      const marks = node.marks || [];
      const link = marks.find((m) => m.type === 'link');
      const run = new TextRun({ ...runStyle(marks), style: link ? 'Hyperlink' : undefined, text: node.text });
      out.push(link ? new ExternalHyperlink({ link: link.attrs.href, children: [run] }) : run);
    } else if (node.type === 'hardBreak') {
      out.push(new TextRun({ break: 1 }));
    } else if (node.type === 'image') {
      const run = await convertImage(node.attrs || {}, ctx);
      if (run) out.push(run);
    }
  }
  return out;
}

function runStyle(marks) {
  const style = {};
  for (const mark of marks) {
    const a = mark.attrs || {};
    switch (mark.type) {
      case 'bold':
        style.bold = true;
        break;
      case 'italic':
        style.italics = true;
        break;
      case 'underline':
        style.underline = {};
        break;
      case 'strike':
        style.strike = true;
        break;
      case 'superscript':
        style.superScript = true;
        break;
      case 'subscript':
        style.subScript = true;
        break;
      case 'code':
        style.font = 'Courier New';
        break;
      case 'highlight':
        style.highlight = 'yellow';
        break;
      case 'textStyle': {
        if (a.fontFamily) style.font = a.fontFamily.split(',')[0].replace(/["']/g, '').trim();
        const size = toHalfPoints(a.fontSize);
        if (size) style.size = size;
        const color = toHex(a.color);
        if (color) style.color = color;
        break;
      }
    }
  }
  return style;
}

export function toHalfPoints(value) {
  if (!value) return undefined;
  const m = String(value).trim().match(/^([\d.]+)\s*(pt|px)?$/i);
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  const pt = (m[2] || 'pt').toLowerCase() === 'px' ? n * 0.75 : n;
  return Math.round(pt * 2);
}

export function toHex(value) {
  if (!value) return undefined;
  const v = String(value).trim();
  let m = v.match(/^#([0-9a-f]{6})$/i);
  if (m) return m[1].toUpperCase();
  m = v.match(/^#([0-9a-f]{3})$/i);
  if (m) return m[1].split('').map((c) => c + c).join('').toUpperCase();
  m = v.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (m) return m.slice(1, 4).map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase();
  return undefined;
}

async function convertImage(attrs, ctx) {
  const m = String(attrs.src || '').match(/^data:image\/(png|jpe?g|gif|bmp);base64,(.+)$/i);
  if (!m) return null;
  const type = m[1].toLowerCase().startsWith('jp') ? 'jpg' : m[1].toLowerCase();
  const data = base64ToBytes(m[2]);

  let width = Number(attrs.width) || 0;
  let height = Number(attrs.height) || 0;
  if ((!width || !height) && ctx.getImageSize) {
    const size = await ctx.getImageSize(attrs.src);
    width = size.width;
    height = size.height;
  }
  if (!width || !height) {
    width = 400;
    height = 300;
  }
  if (width > CONTENT_WIDTH_PX) {
    height = Math.round((height * CONTENT_WIDTH_PX) / width);
    width = CONTENT_WIDTH_PX;
  }

  return new ImageRun({
    type,
    data,
    transformation: { width, height },
    altText: attrs.alt ? { name: 'image', description: attrs.alt, title: attrs.title || attrs.alt } : undefined,
  });
}

function base64ToBytes(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}
