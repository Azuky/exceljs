const Stream = require('readable-stream');

const Excel = verquire('exceljs');
const WorksheetWriter = verquire('stream/xlsx/worksheet-writer');
const StreamBuf = verquire('utils/stream-buf');

const createSheet = (name = 'sheet') =>
  new Excel.stream.xlsx.WorkbookWriter({
    stream: new Stream.Writable({
      write: function noop(chunk, encoding, callback) {
        callback();
      },
    }),
  }).addWorksheet(name);

describe('Workbook Writer', () => {
  it('generates valid xml even when there is no data', () =>
    // issue: https://github.com/guyonroche/exceljs/issues/99
    // PR: https://github.com/guyonroche/exceljs/pull/255
    new Promise((resolve, reject) => {
      const mockWorkbook = {
        _openStream() {
          return this.stream;
        },
        stream: new StreamBuf(),
      };
      mockWorkbook.stream.on('finish', () => {
        try {
          const xml = mockWorkbook.stream.read().toString();
          expect(xml).xml.to.be.valid();
          resolve();
        } catch (error) {
          reject(error);
        }
      });

      const writer = new WorksheetWriter({
        id: 1,
        workbook: mockWorkbook,
      });

      writer.commit();
    }));

  describe('destroy', () => {
    it('is not a valid operation', () => {
      expect(() => createSheet().destroy()).to.throw('Invalid Operation: destroy');
    });
  });

  describe('column keys', () => {
    it('gets, sets, deletes and iterates keys', () => {
      const ws = createSheet();
      ws.columns = [{ key: 'id' }, { key: 'name' }];

      expect(ws.getColumnKey('id')).to.equal(ws.getColumn(1));
      expect(ws.getColumn('name')).to.equal(ws.getColumn(2));

      ws.setColumnKey('alias', ws.getColumn(2));
      expect(ws.getColumnKey('alias')).to.equal(ws.getColumn(2));

      const keys = [];
      ws.eachColumnKey((column, key) => keys.push(key));
      expect(keys).to.deep.equal(['id', 'name', 'alias']);

      ws.deleteColumnKey('alias');
      expect(ws.getColumnKey('alias')).to.equal(undefined);
    });
  });

  describe('rows', () => {
    it('iterates only rows with values unless includeEmpty is set', () => {
      const ws = createSheet();
      ws.getCell('A1').value = 1;
      ws.getCell('C3').value = 3;

      const withValues = [];
      ws.eachRow((row, number) => withValues.push(number));
      expect(withValues).to.deep.equal([1, 3]);

      const all = [];
      ws.eachRow({ includeEmpty: true }, (row, number) => all.push(number));
      expect(all).to.deep.equal([1, 2, 3]);
    });

    it('finds uncommitted rows and exposes the last one', () => {
      const ws = createSheet();
      ws.getCell('A1').value = 1;
      ws.getCell('A3').value = 3;

      expect(ws.findRow(1).number).to.equal(1);
      expect(ws.findRow(2)).to.equal(undefined);
      expect(ws.findRow(3).number).to.equal(3);
      expect(ws.findRow(9)).to.equal(undefined);
      expect(ws.lastRow.number).to.equal(3);
    });

    it('has no lastRow before any row exists', () => {
      expect(createSheet().lastRow).to.equal(undefined);
    });

    it('refuses to get a row that has already been committed', () => {
      const ws = createSheet();
      ws.addRow([1]).commit();
      expect(() => ws.getRow(1)).to.throw('Out of bounds: this row has been committed');
    });
  });

  describe('conditional formatting', () => {
    const refs = (ws) => ws.conditionalFormatting.map((cf) => cf.ref);
    const createWithThree = () => {
      const ws = createSheet();
      ['A1', 'B1', 'C1'].forEach((ref) => ws.addConditionalFormatting({ ref, rules: [] }));
      return ws;
    };

    it('removes by index', () => {
      const ws = createWithThree();
      ws.removeConditionalFormatting(0);
      expect(refs(ws)).to.deep.equal(['B1', 'C1']);
    });

    it('keeps only the entries accepted by a filter function', () => {
      const ws = createWithThree();
      ws.removeConditionalFormatting((cf) => cf.ref === 'B1');
      expect(refs(ws)).to.deep.equal(['B1']);
    });

    it('removes everything when called without a filter', () => {
      const ws = createWithThree();
      ws.removeConditionalFormatting();
      expect(refs(ws)).to.deep.equal([]);
    });
  });

  describe('protection', () => {
    it('hashes the password with the given spinCount', async () => {
      const ws = createSheet();
      await ws.protect('secret', { spinCount: 10 });
      expect(ws.sheetProtection).to.include({ sheet: true, algorithmName: 'SHA-512', spinCount: 10 });
      expect(ws.sheetProtection.hashValue).to.be.a('string');
      expect(ws.sheetProtection.saltValue).to.be.a('string');
    });

    it('drops spinCount when there is no password', async () => {
      const ws = createSheet();
      await ws.protect(undefined, { spinCount: 10, objects: true });
      expect(ws.sheetProtection).to.deep.equal({ sheet: true, objects: true });
    });

    it('unprotect clears the protection', async () => {
      const ws = createSheet();
      await ws.protect('secret');
      ws.unprotect();
      expect(ws.sheetProtection).to.equal(null);
    });
  });

  describe('commit', () => {
    it('is a no-op the second time', () => {
      const ws = createSheet();
      ws.addRow([1]);
      ws.commit();
      expect(ws.committed).to.equal(true);
      expect(() => ws.commit()).to.not.throw();
    });
  });
});
