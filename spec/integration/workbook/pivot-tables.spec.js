// *Note*: `fs.promises` not supported before Node.js 11.14.0;
// ExcelJS version range '>=8.3.0' (as of 2023-10-08).
const fs = require('fs');
const { promisify } = require('util');

const fsReadFileAsync = promisify(fs.readFile);

const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const PIVOT_TABLE_FILEPATHS = [
  'xl/pivotCache/pivotCacheRecords1.xml',
  'xl/pivotCache/pivotCacheDefinition1.xml',
  'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
  'xl/pivotTables/pivotTable1.xml',
  'xl/pivotTables/_rels/pivotTable1.xml.rels',
];

const TEST_XLSX_FILEPATH = './spec/out/wb.test.xlsx';

const TEST_DATA = [
  ['A', 'B', 'C', 'D', 'E'],
  ['a1', 'b1', 'c1', 4, 5],
  ['a1', 'b2', 'c1', 4, 5],
  ['a2', 'b1', 'c2', 14, 24],
  ['a2', 'b2', 'c2', 24, 35],
  ['a3', 'b1', 'c3', 34, 45],
  ['a3', 'b2', 'c3', 44, 45],
];

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Pivot Tables', () => {
    it('if pivot table added, then certain xml and rels files are added', async () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');
      worksheet2.addPivotTable({
        sourceSheet: worksheet1,
        rows: ['A', 'B'],
        columns: ['C'],
        values: ['E'],
        metric: 'sum',
      });

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(zip.files[filepath]).to.not.be.undefined;
        }
      });
    });

    it('if pivot table NOT added, then certain xml and rels files are not added', () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      workbook.addWorksheet('Sheet2');

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(zip.files[filepath]).to.be.undefined;
        }
      });
    });

    describe('validation', () => {
      let warn;
      beforeEach(() => {
        warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
      });
      afterEach(() => {
        warn.mockRestore();
      });

      function setup() {
        const workbook = new ExcelJS.Workbook();
        const source = workbook.addWorksheet('Source');
        source.addRows(TEST_DATA);
        const target = workbook.addWorksheet('Target');
        return { workbook, source, target };
      }

      const model = (source, overrides) => ({
        sourceSheet: source,
        rows: ['A', 'B'],
        columns: ['C'],
        values: ['E'],
        metric: 'sum',
        ...overrides,
      });

      it('accepts a valid model and registers the pivot table on sheet and workbook', () => {
        const { workbook, source, target } = setup();
        const pivotTable = target.addPivotTable(model(source));
        expect(target.pivotTables).to.deep.equal([pivotTable]);
        expect(workbook.pivotTables).to.deep.equal([pivotTable]);
        expect(pivotTable.metric).to.equal('sum');
        expect(pivotTable.rows).to.deep.equal([0, 1]);
        expect(pivotTable.columns).to.deep.equal([2]);
        expect(pivotTable.values).to.deep.equal([4]);
        expect(warn).toHaveBeenCalledTimes(1);
      });

      it('accepts a model without a metric', () => {
        const { source, target } = setup();
        const { metric, ...withoutMetric } = model(source);
        expect(() => target.addPivotTable(withoutMetric)).not.to.throw();
      });

      it('allows at most one pivot table per workbook', () => {
        const { source, target } = setup();
        target.addPivotTable(model(source));
        expect(() => target.addPivotTable(model(source))).to.throw(
          'A pivot table was already added. At this time, ExcelJS supports at most one pivot table per file.',
        );
      });

      it('supports only the sum metric', () => {
        const { source, target } = setup();
        expect(() => target.addPivotTable(model(source, { metric: 'count' }))).to.throw(
          'Only the "sum" metric is supported at this time.',
        );
      });

      it('rejects a header name that is not in the source sheet', () => {
        const { source, target } = setup();
        expect(() => target.addPivotTable(model(source, { rows: ['A', 'Nope'] }))).to.throw(
          'The header name "Nope" was not found in Source.',
        );
        expect(() => target.addPivotTable(model(source, { columns: ['Z'] }))).to.throw(
          'The header name "Z" was not found in Source.',
        );
        expect(() => target.addPivotTable(model(source, { values: ['Q'] }))).to.throw(
          'The header name "Q" was not found in Source.',
        );
      });

      it('requires at least one row', () => {
        const { source, target } = setup();
        expect(() => target.addPivotTable(model(source, { rows: [] }))).to.throw('No pivot table rows specified.');
      });

      it('requires at least one column', () => {
        const { source, target } = setup();
        expect(() => target.addPivotTable(model(source, { columns: [] }))).to.throw(
          'No pivot table columns specified.',
        );
      });

      it('requires exactly one value', () => {
        const { source, target } = setup();
        expect(() => target.addPivotTable(model(source, { values: [] }))).to.throw(
          'Exactly 1 value needs to be specified at this time.',
        );
        expect(() => target.addPivotTable(model(source, { values: ['D', 'E'] }))).to.throw(
          'Exactly 1 value needs to be specified at this time.',
        );
      });

      it('does not register a pivot table that failed validation', () => {
        const { workbook, source, target } = setup();
        expect(() => target.addPivotTable(model(source, { rows: [] }))).to.throw('No pivot table rows specified.');
        expect(target.pivotTables).to.deep.equal([]);
        expect(workbook.pivotTables).to.deep.equal([]);
      });
    });
  });
});
