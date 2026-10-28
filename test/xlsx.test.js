import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Workbook } from '../src/excel/workbook.js';

const { readXlsx, writeXlsx, addPrefixes, stripPrefixes } = createRequire(import.meta.url)('../electron/xlsx-io.cjs');

function sampleBook() {
  const wb = new Workbook();
  const id = wb.activeId;
  const rows = [
    ['Item', 'Price', 'Qty', 'Total'],
    ['Apples', '$1.25', '4', '=B2*C2'],
    ['Pears', '$2.00', '3', '=B3*C3'],
    ['Due', '3/15/2026', '', '=SUM(D2:D3)'],
    ['Check', '=IF(D4>10,TRUE,FALSE)', '=TEXTJOIN("-",TRUE,A2:A3)', '=1/0'],
  ];
  rows.forEach((row, r) => row.forEach((text, c) => text && wb.enter(id, r, c, text)));
  wb.setStyle(id, { r0: 0, c0: 0, r1: 0, c1: 3 }, { bold: true, fill: '#dce6f4', border: { bottom: 'thin' } });
  wb.setStyle(id, { r0: 1, c0: 3, r1: 3, c1: 3 }, { numFmt: '$#,##0.00', align: 'right' });
  wb.setStyle(id, { r0: 4, c0: 0, r1: 4, c1: 0 }, { italic: true, color: '#c00000', font: 'Arial', size: 14, wrap: true });
  wb.setColWidth(id, [0], 120);
  wb.setRowHeight(id, [0], 30);
  wb.setFrozen(id, 1, 0);
  wb.addSheet();
  wb.enter(wb.activeId, 0, 0, '=Sheet1!D4*2');
  wb.setActive(id);
  return { wb, id };
}

test('saves and reopens values, formulas, formats and layout', async () => {
  const { wb } = sampleBook();
  const model = await readXlsx(await writeXlsx(wb.toModel()));
  const reopened = new Workbook(model);
  const [s1, s2] = reopened.sheetIds;
  assert.equal(reopened.sheetName(s1), 'Sheet1');
  assert.equal(reopened.getValue(s1, 1, 1), 1.25);
  assert.equal(reopened.getValue(s1, 3, 3), 11);
  assert.equal(reopened.getInput(s1, 3, 3), '=SUM(D2:D3)');
  assert.equal(reopened.getValue(s1, 3, 1), 46096); // 3/15/2026
  assert.equal(reopened.getNumFmt(s1, 3, 1), 'm/d/yyyy');
  assert.equal(reopened.getValue(s1, 4, 1), true);
  assert.equal(reopened.getValue(s1, 4, 2), 'Apples-Pears');
  assert.equal(reopened.getValue(s1, 4, 3).error, '#DIV/0!');
  assert.equal(reopened.getValue(s2, 0, 0), 22);
  assert.deepEqual(reopened.getStyle(s1, 0, 0), { bold: true, fill: '#dce6f4', border: { bottom: 'thin' } });
  assert.deepEqual(reopened.getStyle(s1, 2, 3), { numFmt: '$#,##0.00', align: 'right' });
  assert.deepEqual(reopened.getStyle(s1, 4, 0), { italic: true, color: '#c00000', font: 'Arial', size: 14, wrap: true });
  assert.equal(reopened.colWidth(s1, 0), 120);
  assert.equal(reopened.rowHeight(s1, 0), 30);
  assert.deepEqual(reopened.metaOf(s1).frozen, { rows: 1, cols: 0 });
  assert.equal(reopened.activeId, s1);
});

test('saved file stores formula results, so viewers that do not calculate still show them', async () => {
  const { wb } = sampleBook();
  const model = await readXlsx(await writeXlsx(wb.toModel()));
  const total = model.sheets[0].cells.find((c) => c.r === 3 && c.c === 3);
  assert.equal(total.value, 11);
});

test('newer Excel functions get the prefix Excel requires, and lose it on open', () => {
  assert.equal(addPrefixes('TEXTJOIN(",",TRUE,A1)+SUM(B1)'), '_xlfn.TEXTJOIN(",",TRUE,A1)+SUM(B1)');
  assert.equal(addPrefixes('STDEV.S(A1:A3)+FILTER(A1:A3,B1:B3)'), '_xlfn.STDEV.S(A1:A3)+_xlfn._xlws.FILTER(A1:A3,B1:B3)');
  assert.equal(addPrefixes('"IFS(" & \'XOR(\'!A1'), '"IFS(" & \'XOR(\'!A1');
  assert.equal(stripPrefixes('_xlfn.XLOOKUP(1,A:A,B:B)'), 'XLOOKUP(1,A:A,B:B)');
});

test('rejects files that are not spreadsheets', async () => {
  await assert.rejects(readXlsx(Buffer.from('not a zip')));
});

// LibreOffice is the independent check that saved files open elsewhere.
// It runs where LibreOffice is installed and is skipped otherwise.
const soffice = ['soffice', 'libreoffice'].find((cmd) => {
  try {
    execFileSync(cmd, ['--version'], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
});

test('LibreOffice opens a saved file and calculates the same results', { skip: !soffice && 'LibreOffice not installed' }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'my-excel-'));
  const { wb } = sampleBook();
  fs.writeFileSync(path.join(dir, 'book.xlsx'), await writeXlsx(wb.toModel()));
  // Re-save through LibreOffice: if it can't read the file, this fails.
  execFileSync(soffice, [loProfile(dir), '--headless', '--convert-to', 'ods', '--outdir', path.join(dir, 'ods'), path.join(dir, 'book.xlsx')], {
    stdio: 'ignore',
    timeout: 120000,
  });
  execFileSync(soffice, [loProfile(dir), '--headless', '--convert-to', 'xlsx', '--outdir', path.join(dir, 'back'), path.join(dir, 'ods', 'book.ods')], {
    stdio: 'ignore',
    timeout: 120000,
  });
  const model = await readXlsx(fs.readFileSync(path.join(dir, 'back', 'book.xlsx')));
  const cell = (s, r, c) => model.sheets[s].cells.find((x) => x.r === r && x.c === c);
  assert.equal(cell(0, 3, 3).input, '=SUM(D2:D3)');
  assert.equal(cell(0, 3, 3).value, 11); // calculated by LibreOffice
  assert.equal(cell(0, 4, 2).value, 'Apples-Pears');
  assert.equal(cell(1, 0, 0).value, 22);
  assert.equal(cell(0, 0, 0).style.bold, true);
  assert.equal(cell(0, 3, 1).input, 46096);
  // And our app reads LibreOffice's file back.
  const reopened = new Workbook(model);
  assert.equal(reopened.getValue(reopened.sheetIds[0], 3, 3), 11);
  fs.rmSync(dir, { recursive: true, force: true });
});

// Each test file gets its own LibreOffice profile, so test files running at once don't collide.
function loProfile(dir) {
  return `-env:UserInstallation=file://${path.join(dir, 'lo-profile').replace(/\\/g, '/')}`;
}
