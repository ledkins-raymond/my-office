import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Workbook } from '../src/excel/workbook.js';
import { writePivot, refreshPivot, sourceFields, isNumericField, pivotOnSheet } from '../src/excel/pivot.js';

const { readXlsx, writeXlsx } = createRequire(import.meta.url)('../electron/xlsx-io.cjs');

const DATA = [
  ['Region', 'Item', 'Quarter', 'Amount'],
  ['East', 'Apple', 'Q1', 10],
  ['West', 'Apple', 'Q1', 5],
  ['east', 'Pear', 'Q2', 7],
  [null, 'Pear', 'Q1', 3],
  ['West', 'a*b', 'Q2', 1],
];

function dataBook() {
  const cells = [];
  DATA.forEach((row, r) =>
    row.forEach((v, c) => v !== null && cells.push({ r, c, input: typeof v === 'string' ? `'${v}` : v })),
  );
  cells.push({ r: 1, c: 3, input: 10, style: { numFmt: '$#,##0.00' } });
  const wb = new Workbook({ sheets: [{ name: 'Sales Data', cells }] });
  return { wb, data: wb.activeId };
}

// The pivot sheet as a grid of displayed values.
function values(wb, id) {
  const { rows, cols } = wb.usedSize(id);
  return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => wb.getValue(id, r, c)));
}

const def = (extra) => ({ source: { sheet: 'Sales Data', range: 'A1:D6' }, rows: ['Region'], columns: [], values: [{ field: 'Amount', agg: 'sum' }], ...extra });

test('reads field names and guesses which are numeric', () => {
  const { wb, data } = dataBook();
  const range = { r0: 0, c0: 0, r1: 5, c1: 3 };
  assert.deepEqual(sourceFields(wb, data, range).map((f) => f.name), ['Region', 'Item', 'Quarter', 'Amount']);
  assert.equal(isNumericField(wb, data, range, 3), true);
  assert.equal(isNumericField(wb, data, range, 0), false);
});

test('sums by one row field on a new sheet, grouping case-insensitively', () => {
  const { wb } = dataBook();
  const id = writePivot(wb, def());
  assert.equal(wb.sheetName(id), 'PivotTable1');
  assert.equal(wb.activeId, id);
  assert.deepEqual(values(wb, id), [
    ['Region', 'Sum of Amount'],
    ['East', 17],
    ['West', 6],
    ['(blank)', 3],
    ['Grand Total', 26],
  ]);
  assert.equal(wb.getInput(id, 1, 1), "=SUMIFS('Sales Data'!$D$2:$D$6,'Sales Data'!$A$2:$A$6,$A2)");
  assert.equal(wb.getNumFmt(id, 1, 1), '$#,##0.00');
  assert.equal(pivotOnSheet(wb, id).source.range, 'A1:D6');
});

test('rows by columns with count and average, wildcard text matched exactly', () => {
  const { wb } = dataBook();
  const id = writePivot(wb, def({ rows: ['Item'], columns: ['Quarter'], values: [{ field: 'Amount', agg: 'count' }] }));
  assert.deepEqual(values(wb, id), [
    ['Count of Amount', null, null, null],
    ['Item', 'Q1', 'Q2', 'Grand Total'],
    ['a*b', 0, 1, 1],
    ['Apple', 2, 0, 2],
    ['Pear', 1, 1, 2],
    ['Grand Total', 3, 2, 5],
  ].map((row, i) => (i === 0 ? ['Count of Amount', 'Quarter →', null, null] : row)));
  const avg = writePivot(wb, def({ rows: ['Region', 'Item'], values: [{ field: 'Amount', agg: 'average' }] }));
  const grid = values(wb, avg);
  assert.deepEqual(grid[0], ['Region', 'Item', 'Average of Amount']);
  assert.deepEqual(grid.at(-1), ['Grand Total', null, 5.2]);
  assert.deepEqual(grid.find((r) => r[0] === 'East' && r[1] === 'Apple'), ['East', 'Apple', 10]);
});

test('updates when the data changes, and refresh picks up new rows and categories', () => {
  const { wb, data } = dataBook();
  const id = writePivot(wb, def());
  wb.enter(data, 1, 3, '100');
  assert.equal(wb.getValue(id, 1, 1), 107);
  wb.enter(data, 6, 0, 'North');
  wb.enter(data, 6, 3, '4');
  refreshPivot(wb, pivotOnSheet(wb, id));
  assert.deepEqual(values(wb, id).map((r) => r[0]), ['Region', 'East', 'North', 'West', '(blank)', 'Grand Total']);
  assert.equal(values(wb, id).at(-1)[1], 120);
  assert.equal(pivotOnSheet(wb, id).source.range, 'A1:D7');
  assert.equal(wb.sheetIds.length, 2); // refreshed in place
});

test('creating a pivot is one undo step', () => {
  const { wb, data } = dataBook();
  writePivot(wb, def());
  wb.undo();
  assert.equal(wb.sheetIds.length, 1);
  assert.equal(wb.activeId, data);
  assert.equal(wb.pivots.length, 0);
});

test('bad definitions explain themselves and change nothing', () => {
  const { wb } = dataBook();
  assert.throws(() => writePivot(wb, def({ rows: ['Nope'] })), /"Nope" is no longer in the data/);
  assert.throws(() => writePivot(wb, def({ source: { sheet: 'Gone', range: 'A1:D6' } })), /no longer exists/);
  assert.equal(wb.sheetIds.length, 1);
});

test('pivots survive saving, keep working in LibreOffice, and can be refreshed after reopening', async () => {
  const { wb } = dataBook();
  writePivot(wb, def({ columns: ['Quarter'] }));
  const model = await readXlsx(await writeXlsx(wb.toModel()));
  assert.deepEqual(model.sheets.map((s) => s.name), ['Sales Data', 'PivotTable1']); // the definitions sheet stays hidden
  const reopened = new Workbook(model);
  assert.equal(reopened.pivots.length, 1);
  const id = reopened.sheetIds[1];
  assert.equal(values(reopened, id).at(-1).at(-1), 26);
  refreshPivot(reopened, pivotOnSheet(reopened, id));
  assert.equal(values(reopened, id).at(-1).at(-1), 26);

  let soffice;
  for (const cmd of ['soffice', 'libreoffice']) {
    try {
      execFileSync(cmd, ['--version'], { stdio: 'ignore' });
      soffice = cmd;
      break;
    } catch {}
  }
  if (!soffice) return;
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'my-excel-pivot-'));
  fs.writeFileSync(path.join(dir, 'pivot.xlsx'), await writeXlsx(wb.toModel()));
  execFileSync(soffice, [loProfile(dir), '--headless', '--convert-to', 'ods', '--outdir', path.join(dir, 'ods'), path.join(dir, 'pivot.xlsx')], { stdio: 'ignore', timeout: 120000 });
  execFileSync(soffice, [loProfile(dir), '--headless', '--convert-to', 'xlsx', '--outdir', path.join(dir, 'back'), path.join(dir, 'ods', 'pivot.ods')], { stdio: 'ignore', timeout: 120000 });
  const lo = await readXlsx(fs.readFileSync(path.join(dir, 'back', 'pivot.xlsx')));
  const sheet = lo.sheets.find((s) => s.name === 'PivotTable1');
  const cell = (r, c) => sheet.cells.find((x) => x.r === r && x.c === c);
  assert.equal(cell(2, 1).value, 10); // East, Q1, calculated by LibreOffice
  assert.equal(cell(5, 3).value, 26); // grand total
  fs.rmSync(dir, { recursive: true, force: true });
});

// Each test file gets its own LibreOffice profile, so test files running at once don't collide.
function loProfile(dir) {
  return `-env:UserInstallation=file://${path.join(dir, 'lo-profile').replace(/\\/g, '/')}`;
}
