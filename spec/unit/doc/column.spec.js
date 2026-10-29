const { createSheetMock } = require('../../utils/index');

const Column = verquire('doc/column');

describe('Column', () => {
  it('creates by defn', () => {
    const sheet = createSheetMock();

    sheet.addColumn(1, {
      header: 'Col 1',
      key: 'id1',
      width: 10,
    });

    expect(sheet.getColumn(1).header).to.equal('Col 1');
    expect(sheet.getColumn(1).headers).to.deep.equal(['Col 1']);
    expect(sheet.getCell(1, 1).value).to.equal('Col 1');
    expect(sheet.getColumn('id1')).to.equal(sheet.getColumn(1));

    sheet.getRow(2).values = { id1: 'Hello, World!' };
    expect(sheet.getCell(2, 1).value).to.equal('Hello, World!');
  });

  it('maintains properties', () => {
    const sheet = createSheetMock();

    const column = sheet.addColumn(1);

    column.key = 'id1';
    expect(sheet._keys.id1).to.equal(column);

    expect(column.number).to.equal(1);
    expect(column.letter).to.equal('A');

    column.header = 'Col 1';
    expect(sheet.getColumn(1).header).to.equal('Col 1');
    expect(sheet.getColumn(1).headers).to.deep.equal(['Col 1']);
    expect(sheet.getCell(1, 1).value).to.equal('Col 1');

    column.header = ['Col A1', 'Col A2'];
    expect(sheet.getColumn(1).header).to.deep.equal(['Col A1', 'Col A2']);
    expect(sheet.getColumn(1).headers).to.deep.equal(['Col A1', 'Col A2']);
    expect(sheet.getCell(1, 1).value).to.equal('Col A1');
    expect(sheet.getCell(2, 1).value).to.equal('Col A2');

    sheet.getRow(3).values = { id1: 'Hello, World!' };
    expect(sheet.getCell(3, 1).value).to.equal('Hello, World!');
  });

  it('creates model', () => {
    const sheet = createSheetMock();

    sheet.addColumn(1, {
      header: 'Col 1',
      key: 'id1',
      width: 10,
    });
    sheet.addColumn(2, {
      header: 'Col 2',
      key: 'name',
      width: 10,
    });
    sheet.addColumn(3, {
      header: 'Col 2',
      key: 'dob',
      width: 10,
      outlineLevel: 1,
    });

    const model = Column.toModel(sheet.columns);
    expect(model.length).to.equal(2);

    expect(model[0].width).to.equal(10);
    expect(model[0].outlineLevel).to.equal(0);
    expect(model[0].collapsed).to.equal(false);

    expect(model[1].width).to.equal(10);
    expect(model[1].outlineLevel).to.equal(1);
    expect(model[1].collapsed).to.equal(true);
  });

  it('gets column values', () => {
    const sheet = createSheetMock();
    sheet.getCell(1, 1).value = 'a';
    sheet.getCell(2, 1).value = 'b';
    sheet.getCell(4, 1).value = 'd';

    expect(sheet.getColumn(1).values).to.deep.equal([, 'a', 'b', , 'd']);
  });
  it('sets column values', () => {
    const sheet = createSheetMock();

    sheet.getColumn(1).values = [2, 3, 5, 7, 11];

    expect(sheet.getCell(1, 1).value).to.equal(2);
    expect(sheet.getCell(2, 1).value).to.equal(3);
    expect(sheet.getCell(3, 1).value).to.equal(5);
    expect(sheet.getCell(4, 1).value).to.equal(7);
    expect(sheet.getCell(5, 1).value).to.equal(11);
    expect(sheet.getCell(6, 1).value).to.equal(null);
  });
  it('sets sparse column values', () => {
    const sheet = createSheetMock();
    const values = [];
    values[2] = 2;
    values[3] = 3;
    values[5] = 5;
    values[11] = 11;
    sheet.getColumn(1).values = values;

    expect(sheet.getCell(1, 1).value).to.equal(null);
    expect(sheet.getCell(2, 1).value).to.equal(2);
    expect(sheet.getCell(3, 1).value).to.equal(3);
    expect(sheet.getCell(4, 1).value).to.equal(null);
    expect(sheet.getCell(5, 1).value).to.equal(5);
    expect(sheet.getCell(6, 1).value).to.equal(null);
    expect(sheet.getCell(7, 1).value).to.equal(null);
    expect(sheet.getCell(8, 1).value).to.equal(null);
    expect(sheet.getCell(9, 1).value).to.equal(null);
    expect(sheet.getCell(10, 1).value).to.equal(null);
    expect(sheet.getCell(11, 1).value).to.equal(11);
    expect(sheet.getCell(12, 1).value).to.equal(null);
  });
  it('sets sparse column values', () => {
    const sheet = createSheetMock();
    sheet.getColumn(1).values = [, , 2, 3, , 5, , 7, , , , 11];

    expect(sheet.getCell(1, 1).value).to.equal(null);
    expect(sheet.getCell(2, 1).value).to.equal(2);
    expect(sheet.getCell(3, 1).value).to.equal(3);
    expect(sheet.getCell(4, 1).value).to.equal(null);
    expect(sheet.getCell(5, 1).value).to.equal(5);
    expect(sheet.getCell(6, 1).value).to.equal(null);
    expect(sheet.getCell(7, 1).value).to.equal(7);
    expect(sheet.getCell(8, 1).value).to.equal(null);
    expect(sheet.getCell(9, 1).value).to.equal(null);
    expect(sheet.getCell(10, 1).value).to.equal(null);
    expect(sheet.getCell(11, 1).value).to.equal(11);
    expect(sheet.getCell(12, 1).value).to.equal(null);
  });
  it('sets default column width', () => {
    const sheet = createSheetMock();

    sheet.addColumn(1, {
      header: 'Col 1',
      key: 'id1',
      style: {
        numFmt: '0.00%',
      },
    });
    sheet.addColumn(2, {
      header: 'Col 2',
      key: 'id2',
      style: {
        numFmt: '0.00%',
      },
      width: 10,
    });
    sheet.getColumn(3).numFmt = '0.00%';

    const model = Column.toModel(sheet.columns);
    expect(model.length).to.equal(3);

    expect(model[0].width).to.equal(9);

    expect(model[1].width).to.equal(10);

    expect(model[2].width).to.equal(9);
  });

  describe('collapsed', () => {
    it('is true once the outline level reaches the sheet outlineLevelCol', () => {
      const sheet = createSheetMock();
      sheet.properties.outlineLevelCol = 2;
      const column = sheet.getColumn(1);

      column.outlineLevel = 1;
      expect(column.collapsed).to.equal(false);
      column.outlineLevel = 2;
      expect(column.collapsed).to.equal(true);
      column.outlineLevel = 3;
      expect(column.collapsed).to.equal(true);
    });

    it('is false for a column without an outline level', () => {
      const sheet = createSheetMock();
      sheet.properties.outlineLevelCol = 0;
      const column = sheet.getColumn(1);
      expect(column.outlineLevel).to.equal(0);
      expect(column.collapsed).to.equal(false);
    });
  });

  describe('styles', () => {
    it('reads back font, alignment, border, fill, protection and numFmt it was given', () => {
      const sheet = createSheetMock();
      const column = sheet.getColumn(1);

      column.font = { bold: true };
      column.alignment = { vertical: 'top' };
      column.border = { left: { style: 'thin' } };
      column.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } };
      column.protection = { locked: false };
      column.numFmt = '0.00%';

      expect(column.font).to.deep.equal({ bold: true });
      expect(column.alignment).to.deep.equal({ vertical: 'top' });
      expect(column.border).to.deep.equal({ left: { style: 'thin' } });
      expect(column.fill.fgColor).to.deep.equal({ argb: 'FF00FF00' });
      expect(column.protection).to.deep.equal({ locked: false });
      expect(column.numFmt).to.equal('0.00%');
    });

    it('applies a style to the cells already in the column', () => {
      const sheet = createSheetMock();
      const column = sheet.getColumn(1);
      sheet.getCell(1, 1).value = 'a';
      sheet.getCell(2, 1).value = 'b';

      column.protection = { locked: false };

      expect(sheet.getCell(1, 1).protection).to.deep.equal({ locked: false });
      expect(sheet.getCell(2, 1).protection).to.deep.equal({ locked: false });
    });
  });

  describe('toString', () => {
    it('describes key, width and headers as JSON', () => {
      const sheet = createSheetMock();
      const column = sheet.addColumn(1, { header: 'Col 1', key: 'id1', width: 10 });
      expect(JSON.parse(column.toString())).to.deep.equal({ key: 'id1', width: 10, headers: ['Col 1'] });
    });
  });
});
