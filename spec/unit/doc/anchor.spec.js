const { knownBug } = require('../../utils/known-bug');

const { createSheetMock } = require('../../utils/index');

const Anchor = verquire('doc/anchor');

describe('Anchor', () => {
  describe('colWidth', () => {
    it('should colWidth equals 640000 when worksheet is undefined', () => {
      const anchor = new Anchor();
      expect(anchor.colWidth).to.equal(640000);
    });
    it('should colWidth equals 640000 when column has not set custom width', () => {
      const anchor = new Anchor(createSheetMock());
      expect(anchor.colWidth).to.equal(640000);
    });
    it('should colWidth equals column width', () => {
      const worksheet = createSheetMock();
      const anchor = new Anchor(worksheet);
      worksheet.addColumn(anchor.nativeCol + 1, {
        width: 10,
      });
      expect(anchor.colWidth).to.equal(worksheet.getColumn(anchor.nativeCol + 1).width * 10000);
    });
  });
  describe('rowHeight', () => {
    it('should rowHeight equals 180000 when worksheet is undefined', () => {
      const anchor = new Anchor();
      expect(anchor.rowHeight).to.equal(180000);
    });
    it('should rowHeight equals 180000 when row has not set height', () => {
      const anchor = new Anchor(createSheetMock());
      expect(anchor.rowHeight).to.equal(180000);
    });
    it('should rowHeight equals row height', () => {
      const worksheet = createSheetMock();
      worksheet.getRow(1).height = 10;

      const anchor = new Anchor(worksheet);
      expect(anchor.rowHeight).to.equal(worksheet.getRow(1).height * 10000);
    });
  });
  describe('resize worksheet`s cells', () => {
    const context = this;
    beforeAll(() => {
      context.worksheet = createSheetMock();
      context.worksheet.getColumn(1).width = 20;
      context.worksheet.getRow(1).height = 20;

      context.anchor = new Anchor(context.worksheet, { col: 0.6, row: 0.6 });
    });

    it('should update colWidth', () => {
      const pre = context.anchor.colWidth;
      context.worksheet.getColumn(1).width *= 2;
      expect(context.anchor.colWidth).to.not.equal(pre);
      expect(context.anchor.colWidth).to.equal(pre * 2);
    });
    it('should update rowHeight', () => {
      const pre = context.anchor.rowHeight;
      context.worksheet.getRow(1).height *= 2;
      expect(context.anchor.rowHeight).to.not.equal(pre);
      expect(context.anchor.rowHeight).to.equal(pre * 2);
    });
    it('should recalculate col', () => {
      const pre = context.anchor.col;
      context.worksheet.getColumn(1).width *= 2;
      expect(context.anchor.col).to.not.equal(pre);
    });
    it('should recalculate row', () => {
      const pre = context.anchor.row;
      context.worksheet.getRow(1).height *= 2;
      expect(context.anchor.row).to.not.equal(pre);
    });
    it('should integer part of row and rowOff should be always equals', () => {
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getRow(1).height *= 2;
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getRow(1).height /= 4;
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getRow(1).height = 0.1;
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getRow(1).height = 9999;
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
    });
    it('should integer part of col and colOff should be always equals', () => {
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getColumn(1).width *= 2;
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getColumn(1).width /= 4;
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getColumn(1).width = 0.1;
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getColumn(1).width = 9999;
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
    });
    it('should update nativeColOff after col has been changed', () => {
      const pre = context.anchor.nativeColOff;
      context.anchor.col -= 0.321;
      expect(context.anchor.nativeColOff).to.not.equal(pre);
    });
    it('should update nativeRowOff after row has been changed', () => {
      const pre = context.anchor.nativeRowOff;
      context.anchor.row -= 0.321;
      expect(context.anchor.nativeColOff).to.not.equal(pre);
    });
  });

  describe('construction', () => {
    it('starts at the origin without an address', () => {
      expect(new Anchor(createSheetMock()).model).to.deep.equal({
        nativeCol: 0,
        nativeColOff: 0,
        nativeRow: 0,
        nativeRowOff: 0,
      });
    });

    it('decodes a string address, applying the offset to the 1-based column and row', () => {
      expect(new Anchor(createSheetMock(), 'B3', -1).model).to.deep.equal({
        nativeCol: 1,
        nativeColOff: 0,
        nativeRow: 2,
        nativeRowOff: 0,
      });
      expect(new Anchor(createSheetMock(), 'B3').model).to.deep.equal({
        nativeCol: 2,
        nativeColOff: 0,
        nativeRow: 3,
        nativeRowOff: 0,
      });
    });

    it('accepts a native model, defaulting missing parts to 0', () => {
      const anchor = new Anchor(createSheetMock(), { nativeCol: 4, nativeRowOff: 7 });
      expect(anchor.model).to.deep.equal({ nativeCol: 4, nativeColOff: 0, nativeRow: 0, nativeRowOff: 7 });
    });

    it('accepts fractional col and row, splitting them into cell and offset', () => {
      const anchor = new Anchor(createSheetMock(), { col: 1.5, row: 2.25 });
      expect(anchor.nativeCol).to.equal(1);
      expect(anchor.nativeColOff).to.equal(320000);
      expect(anchor.nativeRow).to.equal(2);
      expect(anchor.nativeRowOff).to.equal(45000);
      expect(anchor.col).to.be.closeTo(1.5, 0.0001);
      expect(anchor.row).to.be.closeTo(2.25, 0.0001);
    });

    it('falls back to the origin for an unrecognised object', () => {
      expect(new Anchor(createSheetMock(), {}).model).to.deep.equal({
        nativeCol: 0,
        nativeColOff: 0,
        nativeRow: 0,
        nativeRowOff: 0,
      });
    });

    it('asInstance keeps instances and nullish values', () => {
      const anchor = new Anchor(createSheetMock(), 'A1', -1);
      expect(Anchor.asInstance(anchor)).to.equal(anchor);
      expect(Anchor.asInstance(null)).to.equal(null);
      expect(Anchor.asInstance(undefined)).to.equal(undefined);
    });

    // known bug: asInstance calls `new Anchor(model)` so the model lands in the worksheet parameter and is ignored — see spec/KNOWN-BUGS.md
    knownBug(
      'asInstance wraps a plain model into an Anchor with the same position',
      () => {
        const wrapped = Anchor.asInstance({ nativeCol: 3, nativeRow: 5 });
        expect(wrapped).to.be.an.instanceOf(Anchor);
        expect(wrapped.model).to.deep.equal({ nativeCol: 3, nativeColOff: 0, nativeRow: 5, nativeRowOff: 0 });
      },
      /nativeCol: \+0.*to deeply equal/s,
    );
  });

  describe('model', () => {
    it('round trips through the model setter', () => {
      const source = new Anchor(createSheetMock(), {
        nativeCol: 2,
        nativeColOff: 100,
        nativeRow: 3,
        nativeRowOff: 200,
      });
      const target = new Anchor(createSheetMock());
      target.model = source.model;
      expect(target.model).to.deep.equal({ nativeCol: 2, nativeColOff: 100, nativeRow: 3, nativeRowOff: 200 });
    });
  });
});
