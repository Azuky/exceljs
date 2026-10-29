import assert from 'node:assert';
import { createRequire } from 'node:module';
import * as namespace from 'exceljs';
import ExcelJS, { Workbook, stream, ValueType } from 'exceljs';

const cjs = createRequire(import.meta.url)('exceljs');

// esm named equals default
assert.ok(Workbook === ExcelJS.Workbook && stream === ExcelJS.stream && ValueType === ExcelJS.ValueType);

// esm named exports match the cjs keys (no drift)
assert.deepStrictEqual(
  Object.keys(namespace)
    .filter((k) => k !== 'default')
    .sort(),
  Object.keys(cjs).sort(),
);

// require and import share the same classes (no dual-package hazard)
assert.strictEqual(cjs.Workbook, Workbook);

// writeBuffer round-trip
const wb = new Workbook();
wb.addWorksheet('Sheet1').getCell('A1').value = 7;
const buffer = await wb.xlsx.writeBuffer();
const loaded = new Workbook();
await loaded.xlsx.load(buffer);
assert.strictEqual(loaded.getWorksheet('Sheet1').getCell('A1').value, 7);

console.log('consumer.mjs OK');
