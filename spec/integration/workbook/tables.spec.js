const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Tables', () => {
    it('serialises a table per worksheet into uniquely named parts', async () => {
      const wb = new ExcelJS.Workbook();
      const ws1 = wb.addWorksheet('first');
      ws1.addTable({
        name: 'FirstTable',
        ref: 'A1',
        columns: [{ name: 'Id' }, { name: 'Word' }],
        rows: [
          [1, 'Bird'],
          [2, 'is'],
        ],
      });
      const ws2 = wb.addWorksheet('second');
      ws2.addTable({
        name: 'SecondTable',
        ref: 'B2',
        columns: [{ name: 'Key' }, { name: 'Value' }, { name: 'Flag' }],
        rows: [
          ['a', 10, true],
          ['b', 20, false],
          ['c', 30, true],
        ],
      });

      const buffer = await wb.xlsx.writeBuffer();

      const zip = await JSZip.loadAsync(buffer);
      const tableParts = Object.keys(zip.files)
        .filter((name) => name.startsWith('xl/tables/') && !zip.files[name].dir)
        .sort();
      expect(tableParts).to.deep.equal(['xl/tables/table1.xml', 'xl/tables/table2.xml']);

      const table1Xml = await zip.file('xl/tables/table1.xml').async('string');
      const table2Xml = await zip.file('xl/tables/table2.xml').async('string');
      expect(table1Xml).to.match(/<table [^>]*\bid="1"/);
      expect(table1Xml).to.match(/<table [^>]*\bname="FirstTable"/);
      expect(table2Xml).to.match(/<table [^>]*\bid="2"/);
      expect(table2Xml).to.match(/<table [^>]*\bname="SecondTable"/);

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);

      const t1 = wb2.getWorksheet('first').getTables();
      expect(t1.map((t) => t.name)).to.deep.equal(['FirstTable']);
      expect(t1[0].table.tableRef).to.equal('A1:B3');
      expect(t1[0].table.columns.map((c) => c.name)).to.deep.equal(['Id', 'Word']);
      expect(wb2.getWorksheet('first').getRow(1).values).to.deep.equal([undefined, 'Id', 'Word']);
      expect(wb2.getWorksheet('first').getRow(2).values).to.deep.equal([undefined, 1, 'Bird']);
      expect(wb2.getWorksheet('first').getRow(3).values).to.deep.equal([undefined, 2, 'is']);

      const t2 = wb2.getWorksheet('second').getTables();
      expect(t2.map((t) => t.name)).to.deep.equal(['SecondTable']);
      expect(t2[0].table.tableRef).to.equal('B2:D5');
      expect(t2[0].table.columns.map((c) => c.name)).to.deep.equal(['Key', 'Value', 'Flag']);
      expect(wb2.getWorksheet('second').getRow(2).values).to.deep.equal([undefined, undefined, 'Key', 'Value', 'Flag']);
      expect(wb2.getWorksheet('second').getRow(4).values).to.deep.equal([undefined, undefined, 'b', 20, false]);
      expect(wb2.getWorksheet('second').getRow(5).values).to.deep.equal([undefined, undefined, 'c', 30, true]);
    });
  });
});
