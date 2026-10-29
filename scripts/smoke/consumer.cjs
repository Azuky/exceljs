'use strict';

const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const ExcelJS = require('exceljs');

async function main() {
  // cjs keys match lib entry
  assert.deepStrictEqual(Object.keys(ExcelJS).sort(), [
    'DocumentType',
    'ErrorValue',
    'FormulaType',
    'ModelContainer',
    'ReadingOrder',
    'RelationshipType',
    'ValueType',
    'Workbook',
    'stream',
  ]);

  // writeBuffer round-trip
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Sheet1').getCell('A1').value = 7;
  const buffer = await wb.xlsx.writeBuffer();
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(buffer);
  assert.strictEqual(loaded.getWorksheet('Sheet1').getCell('A1').value, 7);

  // stream writer round-trip
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'exceljs-smoke-cjs-'));
  try {
    const filename = path.join(dir, 's.xlsx');
    const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
    writer.addWorksheet('S').addRow(['a', 42]).commit();
    await writer.commit();
    const back = new ExcelJS.Workbook();
    await back.xlsx.readFile(filename);
    const row = back.getWorksheet('S').getRow(1);
    assert.strictEqual(row.getCell(1).value, 'a');
    assert.strictEqual(row.getCell(2).value, 42);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
  console.log('consumer.cjs OK');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
