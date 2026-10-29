const { slideFormula } = verquire('utils/shared-formula');

describe('shared-formula', () => {
  describe('slideFormula', () => {
    const expectations = [
      { args: ['A1+1', 'A2', 'A3'], result: 'A2+1' },
      { args: ['A1+1', 'A2', 'B2'], result: 'B1+1' },
      { args: ['SUM(A1:A10)', 'A11', 'B11'], result: 'SUM(B1:B10)' },
      { args: ['$A$1+A1', 'A2', 'A3'], result: '$A$1+A2' },
      { args: ['$A$1+A1', 'A2', 'B2'], result: '$A$1+B1' },
      { args: ['$A1+A1', 'A2', 'A3'], result: '$A2+A2' },
      { args: ['$A1+A1', 'A2', 'B2'], result: '$A1+B1' },
      { args: ['A$1+A1', 'A2', 'A3'], result: 'A$1+A2' },
      { args: ['A$1+A1', 'A2', 'B2'], result: 'B$1+B1' },
    ];
    expectations.forEach(({ args, result }) => {
      it(`${args[0]} from ${args[1]} to ${args[2]}`, () => {
        expect(slideFormula(...args)).to.equal(result);
      });
    });
  });

  describe('slideFormula tokens that are not cell references', () => {
    it('leaves a column beyond XFD alone, treating it as a named range', () => {
      expect(slideFormula('XFE1+A1', 'A1', 'A2')).to.equal('XFE1+A2');
    });

    it('leaves a column longer than three letters alone', () => {
      expect(slideFormula('ABCD1+A1', 'A1', 'A2')).to.equal('ABCD1+A2');
    });

    it('still slides the last valid column XFD', () => {
      expect(slideFormula('XFD1', 'A1', 'A2')).to.equal('XFD2');
    });

    it('slides sheet-qualified references and keeps the sheet name', () => {
      expect(slideFormula('Sheet1!A1', 'A1', 'B2')).to.equal('Sheet1!B2');
    });

    it('does not treat function names followed by a parenthesis as references', () => {
      expect(slideFormula('LOG10(A1)', 'A1', 'A2')).to.equal('LOG10(A2)');
    });

    it('ignores tokens that are not references', () => {
      expect(slideFormula('IF(TRUE,A1,"x")', 'A1', 'A2')).to.equal('IF(TRUE,A2,"x")');
    });
  });
});
