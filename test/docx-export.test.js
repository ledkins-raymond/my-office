import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Packer } from 'docx';
import mammoth from 'mammoth';
import JSZip from 'jszip';
import { buildDocx, toHalfPoints, toHex } from '../src/word/docx-export.js';

// 1x1 red PNG
const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

const sample = {
  type: 'doc',
  content: [
    { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'My Title' }] },
    {
      type: 'paragraph',
      attrs: { textAlign: 'center' },
      content: [
        { type: 'text', text: 'Plain ' },
        { type: 'text', text: 'bold', marks: [{ type: 'bold' }] },
        { type: 'text', text: ' ' },
        { type: 'text', text: 'italic', marks: [{ type: 'italic' }] },
        { type: 'text', text: ' ' },
        { type: 'text', text: 'underlined', marks: [{ type: 'underline' }] },
        { type: 'text', text: ' ' },
        {
          type: 'text',
          text: 'red Arial 14',
          marks: [{ type: 'textStyle', attrs: { color: '#ff0000', fontFamily: 'Arial', fontSize: '14pt' } }],
        },
        { type: 'hardBreak' },
        { type: 'text', text: 'link', marks: [{ type: 'link', attrs: { href: 'https://example.com' } }] },
      ],
    },
    {
      type: 'bulletList',
      content: [
        { type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Bullet one' }] }] },
        {
          type: 'listItem',
          content: [
            { type: 'paragraph', content: [{ type: 'text', text: 'Bullet two' }] },
            {
              type: 'orderedList',
              content: [
                { type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Nested number' }] }] },
              ],
            },
          ],
        },
      ],
    },
    {
      type: 'table',
      content: [
        {
          type: 'tableRow',
          content: [
            { type: 'tableHeader', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Name' }] }] },
            { type: 'tableHeader', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Qty' }] }] },
          ],
        },
        {
          type: 'tableRow',
          content: [
            { type: 'tableCell', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Apples' }] }] },
            { type: 'tableCell', content: [{ type: 'paragraph' }] },
          ],
        },
      ],
    },
    { type: 'pageBreak' },
    { type: 'image', attrs: { src: PNG } },
    { type: 'paragraph' },
  ],
};

async function render(doc) {
  const buffer = await Packer.toBuffer(await buildDocx(doc, { getImageSize: async () => ({ width: 1, height: 1 }) }));
  return buffer;
}

test('round-trips text, formatting, lists, tables and images through .docx', async () => {
  const buffer = await render(sample);
  const { value: html } = await mammoth.convertToHtml(
    { buffer },
    { styleMap: ['u => u'] },
  );
  assert.match(html, /<h1>My Title<\/h1>/);
  assert.match(html, /<strong>bold<\/strong>/);
  assert.match(html, /<em>italic<\/em>/);
  assert.match(html, /<u>underlined<\/u>/);
  assert.match(html, /<a href="https:\/\/example.com">link<\/a>/);
  assert.match(html, /<ul><li>Bullet one<\/li><li>Bullet two<ol><li>Nested number<\/li><\/ol><\/li><\/ul>/);
  assert.match(html, /<table>.*Name.*Qty.*Apples.*<\/table>/);
  assert.match(html, /<img src="data:image\/png;base64,/);
});

test('writes font, size, color and alignment into document.xml', async () => {
  const zip = await JSZip.loadAsync(await render(sample));
  const xml = await zip.file('word/document.xml').async('string');
  assert.match(xml, /<w:jc w:val="center"\/>/);
  assert.match(xml, /w:ascii="Arial"/);
  assert.match(xml, /<w:color w:val="FF0000"\/>/);
  assert.match(xml, /<w:sz w:val="28"\/>/);
  assert.match(xml, /<w:br w:type="page"\/>/);
});

test('an empty document still produces a valid file', async () => {
  const { value } = await mammoth.extractRawText({ buffer: await render({ type: 'doc', content: [] }) });
  assert.equal(value.trim(), '');
});

test('unit helpers', () => {
  assert.equal(toHalfPoints('12pt'), 24);
  assert.equal(toHalfPoints('16px'), 24);
  assert.equal(toHalfPoints(undefined), undefined);
  assert.equal(toHex('#abc'), 'AABBCC');
  assert.equal(toHex('rgb(255, 0, 16)'), 'FF0010');
});
