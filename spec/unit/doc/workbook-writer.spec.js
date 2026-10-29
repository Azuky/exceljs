const Stream = require('readable-stream');

const { knownBug } = require('../../utils/known-bug');

const Excel = verquire('exceljs');

const createWriter = () =>
  new Excel.stream.xlsx.WorkbookWriter({
    stream: new Stream.Writable({
      write: function noop() {},
    }),
  });

describe('Workbook Writer', () => {
  it('returns undefined for non-existant sheet', () => {
    const wb = createWriter();
    wb.addWorksheet('first');
    expect(wb.getWorksheet('w00t')).to.equal(undefined);
  });

  describe('getWorksheet', () => {
    it('finds a sheet by numeric id and by name', () => {
      const wb = createWriter();
      const first = wb.addWorksheet('first');
      const second = wb.addWorksheet('second');

      expect(first.id).to.equal(1);
      expect(second.id).to.equal(2);
      expect(wb.getWorksheet(1)).to.equal(first);
      expect(wb.getWorksheet(2)).to.equal(second);
      expect(wb.getWorksheet('second')).to.equal(second);
      expect(wb.getWorksheet(3)).to.equal(undefined);
    });

    it('returns undefined for an argument that is neither a number nor a string', () => {
      const wb = createWriter();
      wb.addWorksheet('first');
      expect(wb.getWorksheet({})).to.equal(undefined);
    });

    // sheet ids start at 1, so _worksheets[0] is a hole and find(() => true) stops on it
    knownBug(
      'returns the first sheet when called without arguments',
      () => {
        const wb = createWriter();
        const first = wb.addWorksheet('first');
        wb.addWorksheet('second');
        expect(wb.getWorksheet()).to.equal(first);
      },
      /expected undefined to equal/,
    );
  });

  // Workbook#addWorksheet validates names (see Worksheet#name); the streaming writer should apply the same rules
  describe('addWorksheet name validation', () => {
    it('accepts a valid name', () => {
      const wb = createWriter();
      expect(wb.addWorksheet('Hello, World!').name).to.equal('Hello, World!');
    });

    knownBug(
      'rejects names with illegal characters',
      () => {
        const wb = createWriter();
        expect(() => wb.addWorksheet('a/b')).to.throw(
          'Worksheet name a/b cannot include any of the following characters: * ? : \\ / [ ]',
        );
      },
      /expected \[Function\] to throw an error/,
    );

    knownBug(
      'rejects a duplicate name, ignoring case',
      () => {
        const wb = createWriter();
        wb.addWorksheet('dup');
        expect(() => wb.addWorksheet('DUP')).to.throw('Worksheet name already exists: DUP');
      },
      /expected \[Function\] to throw an error/,
    );
  });
});
