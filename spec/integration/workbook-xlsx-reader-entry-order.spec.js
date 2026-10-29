const fs = require('fs');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const TEST_FILE_NAME = './spec/out/wb.entry-order.xlsx';
const SHEET_NAME = 'Ordered Sheet';

const LEADING_ENTRIES = [
  '[Content_Types].xml',
  '_rels/.rels',
  'xl/_rels/workbook.xml.rels',
  'xl/sharedStrings.xml',
  'xl/styles.xml',
  'xl/worksheets/sheet1.xml',
];
const TRAILING_ENTRY = 'xl/workbook.xml';

// Rebuild the zip so that the workbook rels and shared strings come before the
// worksheet, and xl/workbook.xml comes after it.
const reorderZip = async (buffer) => {
  const source = await JSZip.loadAsync(buffer);
  const names = Object.keys(source.files).filter((name) => !source.files[name].dir);
  expect(names).to.include('xl/sharedStrings.xml');

  const ordered = [
    ...LEADING_ENTRIES,
    ...names.filter((name) => !LEADING_ENTRIES.includes(name) && name !== TRAILING_ENTRY),
    TRAILING_ENTRY,
  ];
  const target = new JSZip();
  for (const name of ordered) {
    target.file(name, await source.files[name].async('nodebuffer'));
  }
  return target.generateAsync({ type: 'nodebuffer' });
};

describe('WorkbookReader', () => {
  describe('zip entry order', () => {
    it('reads a worksheet that precedes xl/workbook.xml in the archive', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet(SHEET_NAME);
      ws.addRow(['alpha', 'beta', 42]);
      ws.addRow(['gamma', 'alpha', 7]);

      const reordered = await reorderZip(await wb.xlsx.writeBuffer());

      // sanity check: the produced entry order is the one under test
      const entries = Object.keys((await JSZip.loadAsync(reordered)).files);
      expect(entries.indexOf('xl/_rels/workbook.xml.rels')).to.be.below(entries.indexOf('xl/worksheets/sheet1.xml'));
      expect(entries.indexOf('xl/sharedStrings.xml')).to.be.below(entries.indexOf('xl/worksheets/sheet1.xml'));
      expect(entries.indexOf('xl/worksheets/sheet1.xml')).to.be.below(entries.indexOf(TRAILING_ENTRY));

      fs.mkdirSync('./spec/out', { recursive: true });
      fs.writeFileSync(TEST_FILE_NAME, reordered);

      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(TEST_FILE_NAME, {
        worksheets: 'emit',
        sharedStrings: 'cache',
        styles: 'ignore',
        hyperlinks: 'ignore',
        entries: 'ignore',
      });

      const sheets = [];
      for await (const worksheetReader of workbookReader) {
        const rows = [];
        for await (const row of worksheetReader) {
          rows.push(row.values);
        }
        sheets.push({ name: worksheetReader.name, rows });
      }

      expect(sheets).to.have.length(1);
      expect(sheets[0].name).to.equal(SHEET_NAME);
      expect(sheets[0].rows).to.deep.equal([
        [undefined, 'alpha', 'beta', 42],
        [undefined, 'gamma', 'alpha', 7],
      ]);
    }, 5000);
  });
});
