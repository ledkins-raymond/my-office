import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { Packer } from 'docx';
import { buildDocx } from '../src/word/docx-export.js';

const { importDocx } = createRequire(import.meta.url)('../electron/docx-import.cjs');

const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

const text = (t, marks) => (marks ? { type: 'text', text: t, marks } : { type: 'text', text: t });
const para = (content, attrs) => ({ type: 'paragraph', ...(attrs && { attrs }), ...(content && { content }) });
const item = (...content) => ({ type: 'listItem', content });

async function roundTrip(doc) {
  const buffer = await Packer.toBuffer(await buildDocx(doc, { getImageSize: async () => ({ width: 20, height: 10 }) }));
  return importDocx(buffer);
}

test('keeps headings, alignment and character formatting', async () => {
  const out = await roundTrip({
    type: 'doc',
    content: [
      { type: 'heading', attrs: { level: 2 }, content: [text('Section')] },
      para(
        [
          text('Bold', [{ type: 'bold' }]),
          text(' plain '),
          text('styled', [
            { type: 'italic' },
            { type: 'underline' },
            { type: 'textStyle', attrs: { fontFamily: 'Georgia', fontSize: '16pt', color: '#c00000' } },
          ]),
          { type: 'hardBreak' },
          text('marked', [{ type: 'highlight' }, { type: 'strike' }]),
          text(' link', [{ type: 'link', attrs: { href: 'https://example.com/' } }]),
        ],
        { textAlign: 'center' },
      ),
    ],
  });

  const [heading, p] = out.content;
  assert.deepEqual(heading, { type: 'heading', attrs: { level: 2 }, content: [text('Section')] });
  assert.equal(p.attrs.textAlign, 'center');
  const byText = Object.fromEntries(p.content.filter((n) => n.text).map((n) => [n.text.trim(), n.marks || []]));
  assert.deepEqual(byText.Bold.map((m) => m.type), ['bold']);
  const styled = byText.styled;
  assert.deepEqual(styled.map((m) => m.type).sort(), ['italic', 'textStyle', 'underline']);
  assert.deepEqual(styled.find((m) => m.type === 'textStyle').attrs, {
    fontFamily: 'Georgia',
    fontSize: '16pt',
    color: '#c00000',
  });
  assert.deepEqual(byText.marked.map((m) => m.type).sort(), ['highlight', 'strike']);
  assert.equal(byText.link.find((m) => m.type === 'link').attrs.href, 'https://example.com/');
  assert.ok(p.content.some((n) => n.type === 'hardBreak'));
});

test('rebuilds nested and separate lists', async () => {
  const out = await roundTrip({
    type: 'doc',
    content: [
      {
        type: 'bulletList',
        content: [
          item(para([text('a')])),
          item(para([text('b')]), { type: 'orderedList', content: [item(para([text('b1')])), item(para([text('b2')]))] }),
          item(para([text('c')])),
        ],
      },
      { type: 'orderedList', content: [item(para([text('one')])), item(para([text('two')]))] },
      { type: 'orderedList', content: [item(para([text('again')]))] },
    ],
  });

  assert.deepEqual(
    out.content.map((b) => b.type),
    ['bulletList', 'orderedList', 'orderedList'],
  );
  const [bullets, first, second] = out.content;
  assert.equal(bullets.content.length, 3);
  assert.equal(bullets.content[1].content[1].type, 'orderedList');
  assert.equal(bullets.content[1].content[1].content.length, 2);
  assert.equal(first.content.length, 2);
  assert.equal(second.content.length, 1);
});

test('keeps tables, merged cells, pictures and page breaks', async () => {
  const cell = (t, attrs) => ({ type: 'tableCell', ...(attrs && { attrs }), content: [para([text(t)])] });
  const out = await roundTrip({
    type: 'doc',
    content: [
      {
        type: 'table',
        content: [
          { type: 'tableRow', content: [cell('wide', { colspan: 2 }), cell('x')] },
          { type: 'tableRow', content: [cell('1'), cell('2'), cell('3')] },
        ],
      },
      { type: 'pageBreak' },
      { type: 'image', attrs: { src: PNG, width: 40, height: 20 } },
      para([text('end')]),
    ],
  });

  const [table, pageBreak, image, end] = out.content;
  assert.equal(table.type, 'table');
  assert.equal(table.content[0].content[0].attrs.colspan, 2);
  assert.equal(table.content[1].content.length, 3);
  assert.equal(pageBreak.type, 'pageBreak');
  assert.equal(image.type, 'image');
  assert.match(image.attrs.src, /^data:image\/png;base64,/);
  assert.deepEqual([image.attrs.width, image.attrs.height], [40, 20]);
  assert.equal(end.content[0].text, 'end');
});

test('rejects files that are not Word documents', async () => {
  await assert.rejects(importDocx(Buffer.from('not a zip')));
});
