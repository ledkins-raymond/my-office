import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, parseRef, rangeName, colName } from '../src/excel/workbook.js';

const range = (text) => parseRef(text);
const at = (text) => {
  const { r0, c0 } = parseRef(text);
  return [r0, c0];
};

function book(rows) {
  const cells = [];
  rows.forEach((row, r) => row.forEach((input, c) => input !== null && cells.push({ r, c, input })));
  const wb = new Workbook({ sheets: [{ name: 'Sheet1', cells }] });
  return { wb, id: wb.activeId };
}

test('addresses', () => {
  assert.equal(colName(0), 'A');
  assert.equal(colName(25), 'Z');
  assert.equal(colName(26), 'AA');
  assert.equal(colName(701), 'ZZ');
  assert.deepEqual(parseRef('b3'), { r0: 2, c0: 1, r1: 2, c1: 1 });
  assert.deepEqual(parseRef('$D$5:B2'), { r0: 1, c0: 1, r1: 4, c1: 3 });
  assert.equal(rangeName(parseRef('A1:C3')), 'A1:C3');
  assert.equal(parseRef('hello'), null);
});

test('formulas calculate, including Excel-style TRUE and FALSE', () => {
  const { wb, id } = book([
    [1, "'a"],
    [2, "'b"],
    ['=SUM(A1:A2)', '=VLOOKUP(2,A1:B2,2,FALSE)'],
  ]);
  assert.equal(wb.getValue(id, ...at('A3')), 3);
  assert.equal(wb.getValue(id, ...at('B3')), 'b');
  wb.enter(id, ...at('A1'), '10');
  assert.equal(wb.getValue(id, ...at('A3')), 12);
  assert.equal(wb.getInput(id, ...at('A3')), '=SUM(A1:A2)');
});

test('typing a date, money or a percentage picks a number format', () => {
  const { wb, id } = book([]);
  wb.enter(id, 0, 0, '1/5/2026');
  wb.enter(id, 1, 0, '$1,234.50');
  wb.enter(id, 2, 0, '15%');
  wb.enter(id, 3, 0, '00123');
  wb.enter(id, 4, 0, "'00123");
  assert.equal(wb.getValue(id, 0, 0), 46027);
  assert.equal(wb.getStyle(id, 0, 0).numFmt, 'm/d/yyyy');
  assert.equal(wb.getValue(id, 1, 0), 1234.5);
  assert.equal(wb.getStyle(id, 1, 0).numFmt, '$#,##0.00');
  assert.equal(wb.getValue(id, 2, 0), 0.15);
  assert.equal(wb.getValue(id, 3, 0), 123);
  assert.equal(wb.getValue(id, 4, 0), '00123');
  const saved = wb.toModel().sheets[0].cells;
  assert.deepEqual(saved.find((c) => c.r === 0).input, 46027);
  assert.deepEqual(saved.find((c) => c.r === 4).input, "'00123");
});

test('undo and redo cover values and formatting together', () => {
  const { wb, id } = book([[1]]);
  wb.enter(id, 0, 0, '5');
  wb.setStyle(id, range('A1'), { bold: true });
  assert.equal(wb.getStyle(id, 0, 0).bold, true);
  wb.undo();
  assert.equal(wb.getStyle(id, 0, 0), undefined);
  assert.equal(wb.getValue(id, 0, 0), 5);
  wb.undo();
  assert.equal(wb.getValue(id, 0, 0), 1);
  wb.redo();
  wb.redo();
  assert.equal(wb.getValue(id, 0, 0), 5);
  assert.equal(wb.getStyle(id, 0, 0).bold, true);
});

test('inserting and deleting rows moves formatting and fixes formulas', () => {
  const { wb, id } = book([[1], [2], ['=A1+A2']]);
  wb.setStyle(id, range('A2'), { fill: '#ffff00' });
  wb.setRowHeight(id, [1], 40);
  wb.shift(id, 'row', 1, 2); // two rows above row 2
  assert.equal(wb.getStyle(id, 3, 0).fill, '#ffff00');
  assert.equal(wb.rowHeight(id, 3), 40);
  assert.equal(wb.getInput(id, 4, 0), '=A1+A4');
  wb.shift(id, 'row', 0, -1);
  assert.equal(wb.getStyle(id, 2, 0).fill, '#ffff00');
  assert.equal(wb.getValue(id, 3, 0).error, '#REF!'); // A1 was deleted
  wb.undo();
  wb.undo();
  assert.equal(wb.getStyle(id, 1, 0).fill, '#ffff00');
  assert.equal(wb.getInput(id, 2, 0), '=A1+A2');
});

test('copy and paste adjusts relative references and copies formatting', () => {
  const { wb, id } = book([[1, 10, '=A1*B1'], [2, 20, null]]);
  wb.setStyle(id, range('C1'), { bold: true });
  wb.paste({ sheet: id, range: range('C1') }, id, { r: 1, c: 2 });
  assert.equal(wb.getInput(id, 1, 2), '=A2*B2');
  assert.equal(wb.getValue(id, 1, 2), 40);
  assert.equal(wb.getStyle(id, 1, 2).bold, true);
  wb.undo();
  assert.equal(wb.getInput(id, 1, 2), '');
});

test('fill continues number series and copies formulas', () => {
  const { wb, id } = book([[1, '=A1*2'], [2, null]]);
  wb.fill(id, range('A1:A2'), range('A1:A5'));
  assert.deepEqual([2, 3, 4].map((r) => wb.getValue(id, r, 0)), [3, 4, 5]);
  wb.fill(id, range('B1'), range('B1:B5'));
  assert.equal(wb.getInput(id, 4, 1), '=A5*2');
  assert.equal(wb.getValue(id, 4, 1), 10);
  wb.undo();
  assert.equal(wb.getInput(id, 4, 1), '');
  wb.undo();
  assert.equal(wb.getInput(id, 4, 0), '');
});

test('sort orders rows by a column and keeps rows together', () => {
  const { wb, id } = book([["'pear", 3], ["'apple", 1], [null, 5], ["'Fig", 2]]);
  wb.sort(id, range('A1:B4'), 0, true);
  assert.deepEqual([0, 1, 2, 3].map((r) => wb.getValue(id, r, 0)), ['apple', 'Fig', 'pear', null]);
  assert.deepEqual([0, 1, 2, 3].map((r) => wb.getValue(id, r, 1)), [1, 2, 3, 5]);
  wb.sort(id, range('A1:B4'), 1, false);
  assert.deepEqual([0, 1, 2, 3].map((r) => wb.getValue(id, r, 1)), [5, 3, 2, 1]);
});

test('sheets can be added, renamed (formulas follow) and removed with undo', () => {
  const { wb, id } = book([['=Sheet2!A1*2']]);
  wb.addSheet();
  const second = wb.activeId;
  assert.equal(wb.sheetName(second), 'Sheet2');
  wb.enter(second, 0, 0, '21');
  assert.equal(wb.getValue(id, 0, 0), 42);
  assert.equal(wb.renameSheet(second, 'Data Sheet'), null);
  assert.equal(wb.getInput(id, 0, 0), "='Data Sheet'!A1*2");
  assert.match(wb.renameSheet(second, 'bad/name'), /cannot contain/);
  wb.removeSheet(second);
  assert.equal(wb.sheetIds.length, 1);
  wb.undo();
  assert.equal(wb.sheetIds.length, 2);
  assert.equal(wb.getValue(id, 0, 0), 42);
});

test('a formula this engine lacks shows the result saved by the other program', () => {
  const wb = new Workbook({
    sheets: [{ name: 'S', cells: [{ r: 0, c: 0, input: '=CUBEVALUE("x")', value: 7 }] }],
  });
  assert.equal(wb.getValue(wb.activeId, 0, 0), 7);
  assert.equal(wb.toModel().sheets[0].cells[0].value, 7);
});
