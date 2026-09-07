import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import { rowsToParts, parseDocument } from '../src/parser';
import { compareParts } from '../src/comparison';
import { csvReport, pdfReport } from '../src/report-export';

const context = { bomName: 'Production BOM.xlsx', jobName: 'Cutting list.doc' };
const header = ['Part Name', 'Quantity', 'Material Type', 'Thickness'];
test('BOM includes only allowed materials, independent of column order and casing', () => {
  const parts = rowsToParts([header, ['A', 1, 'Mild Steel', 2], ['B', 2, ' AISI 304 ', 0.8], ['C', 1, 'Aluminium', 3], ['D', 1, '', 2], ['E', 1, 'mild steel', 1], ['F', 1, 'AISI 304L', 2]], 'bom');
  assert.deepEqual(parts.map(p => p.name), ['A', 'B', 'E']);
  assert.deepEqual(parts.map(p => p.thickness), ['2', '0.8', '1']);
  assert.equal(compareParts([parts[0]], [{ ...parts[0] }])[0].severity, 'match');
});
test('material is never substituted for thickness and absent material column fails clearly', () => {
  assert.throws(() => rowsToParts([['Part Name', 'Quantity'], ['A', 1]], 'bom'), /Material/);
  const parts = rowsToParts([header.slice(0, 3), ['A', 1, 'Mild Steel']], 'bom');
  assert.equal(parts[0].thickness, '');
  assert.equal(compareParts(parts, [{ ...parts[0], thickness: '2' }])[0].severity, 'warning');
});
test('job card is not subject to BOM material filter', () => {
  assert.equal(rowsToParts([['Part Name', 'Cut Qty', 'Thickness'], ['PART-A', 3, 2]], 'jobcard').length, 1);
});
test('CSV quoted fields survive material filtering', async () => {
  const file = new File(['Part Name,Quantity,Material Type,Thickness\r\n"A, bracket",1,Mild Steel,2\r\nB,2,Plastic,3'], 'bom.csv');
  const parts = await parseDocument(file, 'bom');
  assert.equal(parts.length, 1);
  assert.equal(parts[0].name, 'A, bracket');
});
test('CSV and PDF reports produce nonempty files', async () => {
  const parts = rowsToParts([header, ['PART-A', 1, 'Mild Steel', 2]], 'bom');
  const results = compareParts(parts, parts);
  for (const format of ['csv', 'pdf'] as const) {
    const report = format === 'csv' ? csvReport(results) : await pdfReport(results, context);
    assert.ok(report.size > 0);
    if (format === 'pdf') assert.equal(Buffer.from(await report.arrayBuffer()).subarray(0, 5).toString(), '%PDF-');
  }
});
