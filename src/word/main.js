import { Editor, Extension } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { TextStyle, FontFamily, FontSize, Color } from '@tiptap/extension-text-style';
import TextAlign from '@tiptap/extension-text-align';
import Highlight from '@tiptap/extension-highlight';
import Image from '@tiptap/extension-image';
import { Table, TableRow, TableCell, TableHeader } from '@tiptap/extension-table';
import { Packer } from 'docx';
import { buildDocx } from './docx-export.js';
import { PageBreak } from './page-break.js';

const office = window.office;
const $ = (id) => document.getElementById(id);

// Word's alignment shortcuts.
const WordShortcuts = Extension.create({
  name: 'wordShortcuts',
  addKeyboardShortcuts() {
    const align = (value) => () => this.editor.commands.setTextAlign(value);
    return {
      'Mod-l': align('left'),
      'Mod-e': align('center'),
      'Mod-r': align('right'),
      'Mod-j': align('justify'),
    };
  },
});

const editor = new Editor({
  element: $('editor'),
  extensions: [
    StarterKit.configure({
      link: { openOnClick: false, autolink: true, defaultProtocol: 'https' },
    }),
    TextStyle,
    FontFamily,
    FontSize,
    Color,
    Highlight,
    TextAlign.configure({ types: ['heading', 'paragraph'] }),
    Image.configure({
      allowBase64: true,
      resize: {
        enabled: true,
        directions: ['top-left', 'top-right', 'bottom-left', 'bottom-right'],
        minWidth: 24,
        minHeight: 24,
        alwaysPreserveAspectRatio: true,
      },
    }),
    Table.configure({ resizable: false }),
    TableRow,
    TableHeader,
    TableCell,
    PageBreak,
    WordShortcuts,
  ],
  content: '<p></p>',
  autofocus: true,
  editorProps: {
    attributes: { spellcheck: 'true' },
    handlePaste: (_view, event) => insertImageFiles(event.clipboardData?.files),
    handleDrop: (_view, event) => insertImageFiles(event.dataTransfer?.files),
  },
});

// ---------- Dirty tracking ----------
let dirty = false;
function setDirty(value) {
  if (dirty === value) return;
  dirty = value;
  office?.setDirty(value);
}
editor.on('update', () => setDirty(true));

// ---------- File operations (main process does the disk work) ----------
office?.onLoad(({ json, html, simplified, fileName }) => {
  try {
    editor.commands.setContent(json || html || '<p></p>', { emitUpdate: false });
  } catch (err) {
    console.error(err);
    editor.commands.setContent('<p></p>', { emitUpdate: false });
    flash(`Could not display ${fileName}`);
    return;
  }
  editor.commands.focus('start');
  setDirty(false);
  refreshStatus();
  flash(simplified ? `Opened ${fileName} with simplified formatting` : `Opened ${fileName}`);
});

office?.onSaved(({ fileName }) => {
  setDirty(false);
  flash(`Saved ${fileName}`);
});

office?.onCollectDocx(async () => {
  const doc = await buildDocx(editor.getJSON(), { getImageSize });
  return new Uint8Array(await Packer.toArrayBuffer(doc));
});

async function getImageSize(src) {
  const img = new window.Image();
  img.src = src;
  await img.decode();
  return { width: img.naturalWidth, height: img.naturalHeight };
}

// ---------- Commands ----------
const chain = () => editor.chain().focus();

const COMMANDS = {
  undo: () => chain().undo().run(),
  redo: () => chain().redo().run(),
  toggleBold: () => chain().toggleBold().run(),
  toggleItalic: () => chain().toggleItalic().run(),
  toggleUnderline: () => chain().toggleUnderline().run(),
  toggleStrike: () => chain().toggleStrike().run(),
  toggleHighlight: () => chain().toggleHighlight().run(),
  unsetAllMarks: () => chain().unsetAllMarks().run(),
  toggleBulletList: () => chain().toggleBulletList().run(),
  toggleOrderedList: () => chain().toggleOrderedList().run(),
  indent: () => chain().sinkListItem('listItem').run(),
  outdent: () => chain().liftListItem('listItem').run(),
  addRowBefore: () => chain().addRowBefore().run(),
  addRowAfter: () => chain().addRowAfter().run(),
  addColumnBefore: () => chain().addColumnBefore().run(),
  addColumnAfter: () => chain().addColumnAfter().run(),
  mergeOrSplit: () => chain().mergeOrSplit().run(),
  toggleHeaderRow: () => chain().toggleHeaderRow().run(),
  deleteRow: () => chain().deleteRow().run(),
  deleteColumn: () => chain().deleteColumn().run(),
  deleteTable: () => chain().deleteTable().run(),
};

// Actions that need the page (dialogs, pickers) rather than the main process.
const PAGE_ACTIONS = {
  undo: COMMANDS.undo,
  redo: COMMANDS.redo,
  find: openFind,
  'insert-table': insertTable,
  'insert-image': () => $('image-file').click(),
  'insert-link': insertLink,
  'insert-page-break': () => chain().setPageBreak().run(),
  'insert-rule': () => chain().setHorizontalRule().run(),
};

office?.onCommand((name) => PAGE_ACTIONS[name]?.());

document.querySelectorAll('[data-cmd]').forEach((btn) => {
  btn.addEventListener('mousedown', (e) => e.preventDefault()); // keep the text selection
  btn.addEventListener('click', () => COMMANDS[btn.dataset.cmd]?.());
});

document.querySelectorAll('[data-action]').forEach((btn) => {
  btn.addEventListener('mousedown', (e) => e.preventDefault());
  btn.addEventListener('click', () => {
    const name = btn.dataset.action;
    if (PAGE_ACTIONS[name]) PAGE_ACTIONS[name]();
    else office?.action(name);
  });
});

document.querySelectorAll('[data-align]').forEach((btn) => {
  btn.addEventListener('mousedown', (e) => e.preventDefault());
  btn.addEventListener('click', () => chain().setTextAlign(btn.dataset.align).run());
});

// ---------- Style, font, size, color ----------
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
$('size').innerHTML = SIZES.map((s) => `<option value="${s}">${s}</option>`).join('');
$('size').value = '11';

$('style').addEventListener('change', (e) => {
  const v = e.target.value;
  if (v === 'p') chain().setParagraph().run();
  else chain().setHeading({ level: Number(v) }).run();
});

$('font').addEventListener('change', (e) => chain().setFontFamily(e.target.value).run());
$('size').addEventListener('change', (e) => chain().setFontSize(`${e.target.value}pt`).run());

$('color').addEventListener('input', (e) => {
  $('color-swatch').style.borderBottomColor = e.target.value;
  chain().setColor(e.target.value).run();
});
$('color-swatch').style.borderBottomColor = $('color').value;

// ---------- Images ----------
$('image-file').addEventListener('change', async (e) => {
  await insertImageFiles(e.target.files);
  e.target.value = '';
});

function insertImageFiles(files) {
  const images = Array.from(files || []).filter((f) => /^image\/(png|jpeg|gif|bmp)$/.test(f.type));
  if (!images.length) return false;
  images.forEach((file) => {
    const reader = new FileReader();
    reader.onload = () => chain().setImage({ src: reader.result, alt: file.name }).run();
    reader.readAsDataURL(file);
  });
  return true;
}

// ---------- Dialogs (Electron has no window.prompt) ----------
function ask(title, fields) {
  const dialog = $('prompt');
  $('prompt-title').textContent = title;
  $('prompt-fields').innerHTML = '';
  const inputs = fields.map((f) => {
    const label = document.createElement('label');
    label.textContent = f.label;
    const input = document.createElement('input');
    Object.assign(input, { type: f.type || 'text', value: f.value ?? '', required: !!f.required });
    if (f.min != null) input.min = f.min;
    if (f.max != null) input.max = f.max;
    label.append(input);
    $('prompt-fields').append(label);
    return input;
  });
  return new Promise((resolve) => {
    dialog.addEventListener(
      'close',
      () => {
        const values = dialog.returnValue === 'ok' ? inputs.map((i) => i.value) : null;
        // Let the browser finish returning focus before the editor takes it back.
        requestAnimationFrame(() => {
          editor.commands.focus();
          resolve(values);
        });
      },
      { once: true },
    );
    dialog.returnValue = '';
    dialog.showModal();
    inputs[0]?.select();
  });
}

async function insertTable() {
  const answer = await ask('Insert table', [
    { label: 'Columns', type: 'number', value: 3, min: 1, max: 20, required: true },
    { label: 'Rows', type: 'number', value: 3, min: 1, max: 100, required: true },
  ]);
  if (!answer) return;
  const [cols, rows] = answer.map((n) => Math.max(1, parseInt(n, 10) || 1));
  chain().insertTable({ rows, cols, withHeaderRow: false }).run();
}

async function insertLink() {
  const current = editor.getAttributes('link').href || '';
  const answer = await ask('Insert link', [{ label: 'Address', value: current || 'https://' }]);
  if (!answer) return;
  const href = answer[0].trim();
  if (!href || href === 'https://') {
    chain().extendMarkRange('link').unsetLink().run();
    return;
  }
  if (editor.state.selection.empty && !current) {
    chain().insertContent({ type: 'text', text: href, marks: [{ type: 'link', attrs: { href } }] }).run();
  } else {
    chain().extendMarkRange('link').setLink({ href }).run();
  }
}

// ---------- Find and replace ----------
function openFind() {
  $('find-bar').hidden = false;
  const selected = editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to);
  if (selected && !selected.includes('\n')) $('find-text').value = selected;
  $('find-text').select();
}

function closeFind() {
  $('find-bar').hidden = true;
  editor.commands.focus();
}

// Text matches inside each text block, as document positions.
function findMatches(query) {
  const matches = [];
  if (!query) return matches;
  const needle = query.toLowerCase();
  editor.state.doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    // Map each character of the block's text back to a document position.
    let text = '';
    const positions = [];
    node.forEach((child, offset) => {
      if (child.isText) {
        for (let i = 0; i < child.text.length; i++) positions.push(pos + 1 + offset + i);
        text += child.text;
      } else {
        positions.push(-1);
        text += '￼';
      }
    });
    const hay = text.toLowerCase();
    for (let i = hay.indexOf(needle); i !== -1; i = hay.indexOf(needle, i + needle.length)) {
      matches.push({ from: positions[i], to: positions[i + needle.length - 1] + 1 });
    }
    return false;
  });
  return matches;
}

function findNext() {
  const query = $('find-text').value;
  const matches = findMatches(query);
  if (!matches.length) {
    $('find-status').textContent = query ? 'No matches' : '';
    return null;
  }
  const after = editor.state.selection.to;
  const match = matches.find((m) => m.from >= after) || matches[0];
  editor.chain().setTextSelection(match).scrollIntoView().run();
  $('find-status').textContent = `${matches.indexOf(match) + 1} of ${matches.length}`;
  return match;
}

function replaceOne() {
  const query = $('find-text').value;
  const { from, to } = editor.state.selection;
  const selected = editor.state.doc.textBetween(from, to);
  if (query && selected.toLowerCase() === query.toLowerCase()) {
    editor.chain().insertContentAt({ from, to }, $('replace-text').value).run();
  }
  findNext();
}

function replaceAll() {
  const matches = findMatches($('find-text').value);
  if (!matches.length) return;
  const replacement = $('replace-text').value;
  // Replace from the end so earlier positions stay valid.
  const tr = editor.state.tr;
  matches
    .slice()
    .reverse()
    .forEach(({ from, to }) => {
      if (replacement) tr.insertText(replacement, from, to);
      else tr.delete(from, to);
    });
  editor.view.dispatch(tr);
  $('find-status').textContent = `Replaced ${matches.length}`;
}

$('find-next').addEventListener('click', findNext);
$('replace-one').addEventListener('click', replaceOne);
$('replace-all').addEventListener('click', replaceAll);
$('find-close').addEventListener('click', closeFind);
$('find-text').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    findNext();
  }
});
$('find-bar').addEventListener('keydown', (e) => {
  if (e.key === 'Escape') closeFind();
});

// ---------- Toolbar state ----------
function refreshToolbar() {
  document.querySelectorAll('[data-mark]').forEach((btn) => {
    btn.classList.toggle('active', editor.isActive(btn.dataset.mark));
  });
  document.querySelectorAll('[data-node]').forEach((btn) => {
    btn.classList.toggle('active', editor.isActive(btn.dataset.node));
  });
  document.querySelectorAll('[data-align]').forEach((btn) => {
    btn.classList.toggle('active', editor.isActive({ textAlign: btn.dataset.align }));
  });

  const heading = [1, 2, 3].find((level) => editor.isActive('heading', { level }));
  $('style').value = heading ? String(heading) : 'p';

  const style = editor.getAttributes('textStyle');
  const font = (style.fontFamily || 'Calibri').split(',')[0].replace(/["']/g, '').trim();
  ensureOption($('font'), font);
  $('font').value = font;

  const size = style.fontSize ? String(parseFloat(style.fontSize)) : heading ? '' : '11';
  if (size) ensureOption($('size'), size);
  $('size').value = size;

  $('table-bar').hidden = !editor.isActive('table');
}

function ensureOption(select, value) {
  if (!Array.from(select.options).some((o) => o.value === value)) {
    select.add(new Option(value, value));
  }
}

function refreshStatus() {
  const text = editor.state.doc.textBetween(0, editor.state.doc.content.size, ' ', ' ');
  const words = (text.match(/\S+/g) || []).length;
  $('word-count').textContent = `${words.toLocaleString()} word${words === 1 ? '' : 's'}`;
}

let flashTimer;
function flash(message) {
  $('status-message').textContent = message;
  clearTimeout(flashTimer);
  flashTimer = setTimeout(() => ($('status-message').textContent = ''), 5000);
}

editor.on('transaction', refreshToolbar);
editor.on('update', refreshStatus);
refreshToolbar();
refreshStatus();

// Clicking the grey area around the page puts the cursor at the end, like Word.
$('canvas').addEventListener('mousedown', (e) => {
  if (e.target === $('canvas')) {
    e.preventDefault();
    editor.commands.focus('end');
  }
});

// Exposed for automated checks.
window.__editor = editor;
