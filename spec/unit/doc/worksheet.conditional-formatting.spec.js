const Excel = verquire('exceljs');

describe('Worksheet', () => {
  describe('Conditional Formatting', () => {
    function sheetWithThree() {
      const ws = new Excel.Workbook().addWorksheet('blort');
      ['A1:A3', 'B1:B3', 'C1:C3'].forEach((ref) => {
        ws.addConditionalFormatting({
          ref,
          rules: [{ type: 'expression', formulae: ['TRUE'], style: { font: { bold: true } } }],
        });
      });
      return ws;
    }

    const refs = (ws) => ws.conditionalFormattings.map((cf) => cf.ref);

    it('adds formattings in order', () => {
      expect(refs(sheetWithThree())).to.deep.equal(['A1:A3', 'B1:B3', 'C1:C3']);
    });

    it('removes a formatting by index', () => {
      const ws = sheetWithThree();
      ws.removeConditionalFormatting(1);
      expect(refs(ws)).to.deep.equal(['A1:A3', 'C1:C3']);
    });

    it('keeps the formattings matching a filter function', () => {
      const ws = sheetWithThree();
      ws.removeConditionalFormatting((cf) => cf.ref !== 'B1:B3');
      expect(refs(ws)).to.deep.equal(['A1:A3', 'C1:C3']);
    });

    it('removes all formattings when called without an argument', () => {
      const ws = sheetWithThree();
      ws.removeConditionalFormatting();
      expect(ws.conditionalFormattings).to.deep.equal([]);
    });

    it('ignores an index past the end', () => {
      const ws = sheetWithThree();
      ws.removeConditionalFormatting(10);
      expect(refs(ws)).to.deep.equal(['A1:A3', 'B1:B3', 'C1:C3']);
    });
  });
});
