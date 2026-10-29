const utils = verquire('utils/utils');

describe('utils', () => {
  describe('xmlEncode', () => {
    it('encodes xml text', () => {
      expect(utils.xmlEncode('<')).to.equal('&lt;');
      expect(utils.xmlEncode('>')).to.equal('&gt;');
      expect(utils.xmlEncode('&')).to.equal('&amp;');
      expect(utils.xmlEncode('"')).to.equal('&quot;');
      expect(utils.xmlEncode("'")).to.equal('&apos;');

      expect(
        utils.xmlEncode(
          'abc\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x20abc\x7f',
        ),
      ).to.equal('abc abc');

      expect(utils.xmlEncode('<a href="www.whatever.com">Talk to the H&</a>')).to.equal(
        '&lt;a href=&quot;www.whatever.com&quot;&gt;Talk to the H&amp;&lt;/a&gt;',
      );

      expect(utils.xmlEncode('new\x0aline')).to.equal('new\x0aline');
    });
  });
  describe('isDateFmt', () => {
    ['yyyy-mm-dd'].forEach((fmt) => {
      it(`'${fmt}' a date`, () => {
        expect(utils.isDateFmt(fmt)).to.be.true;
      });
    });

    ['', '[Green]#,##0 ;[Red](#,##0)'].forEach((fmt) => {
      it(`'${fmt}' is not a date`, () => {
        expect(utils.isDateFmt(fmt)).to.be.false;
      });
    });
  });

  describe('dateToExcel', () => {
    it('should convert date to excel properly', () => {
      const myDate = new Date(Date.UTC(2017, 11, 15, 17, 0, 0, 0));

      const excelDate = utils.dateToExcel(myDate, false);

      expect(excelDate).to.equal(43084.70833333333);
    });
  });

  describe('excelToDate', () => {
    it('should round to the nearest millisecond when parsing excel date', () => {
      const myDate = new Date(Date.UTC(2017, 11, 15, 17, 0, 0, 0));
      const excelDate = utils.dateToExcel(myDate, false);

      const dateConverted = utils.excelToDate(excelDate, false);

      expect(dateConverted).to.deep.equal(myDate);
    });
    it('should not lost millisecond precision when parsing excel date', () => {
      const myDate = new Date(Date.UTC(2017, 11, 15, 17, 0, 0, 0));
      const excelDate = utils.dateToExcel(myDate, false);

      const dateConverted = utils.excelToDate(excelDate, false);

      expect(dateConverted).to.deep.equal(myDate);
    });
  });

  describe('xmlDecode', () => {
    it('decodes the five predefined entities', () => {
      expect(utils.xmlDecode('&lt;a&amp;b&gt;&apos;&quot;')).to.equal('<a&b>\'"');
    });

    it('leaves unknown entities untouched', () => {
      expect(utils.xmlDecode('&unknown; &nbsp; x')).to.equal('&unknown; &nbsp; x');
    });

    it('decodes only one level, so double-encoded entities stay encoded', () => {
      expect(utils.xmlDecode('&amp;lt;')).to.equal('&lt;');
    });

    it('is the inverse of xmlEncode for text with special characters', () => {
      const text = 'a < b && "c" > \'d\'';
      expect(utils.xmlDecode(utils.xmlEncode(text))).to.equal(text);
    });
  });

  describe('getRelsPath', () => {
    it('builds the rels path of a part', () => {
      expect(utils.getRelsPath('xl/worksheets/sheet1.xml')).to.equal('xl/worksheets/_rels/sheet1.xml.rels');
    });

    it('handles a part at the root', () => {
      expect(utils.getRelsPath('/workbook.xml')).to.equal('/_rels/workbook.xml.rels');
    });
  });

  describe('promiseImmediate', () => {
    it('resolves with the given value on a later tick', async () => {
      let ran = false;
      const promise = utils.promiseImmediate('value');
      promise.then(() => {
        ran = true;
      });
      expect(ran).to.equal(false);
      expect(await promise).to.equal('value');
    });
  });

  describe('toSortedArray', () => {
    it('sorts numbers numerically', () => {
      expect(utils.toSortedArray([10, 9, 1])).to.deep.equal([1, 9, 10]);
    });

    it('sorts other values as strings and accepts any iterable', () => {
      expect(utils.toSortedArray(new Set(['b', 'a', 'c']))).to.deep.equal(['a', 'b', 'c']);
      expect(utils.toSortedArray([10, 9, 'x'])).to.deep.equal([10, 9, 'x']);
    });

    it('does not modify its input', () => {
      const input = [3, 1, 2];
      utils.toSortedArray(input);
      expect(input).to.deep.equal([3, 1, 2]);
    });
  });

  describe('range', () => {
    it('counts up to, but excluding, stop', () => {
      expect([...utils.range(0, 3)]).to.deep.equal([0, 1, 2]);
    });

    it('supports a step and a negative step', () => {
      expect([...utils.range(0, 10, 4)]).to.deep.equal([0, 4, 8]);
      expect([...utils.range(3, 0, -1)]).to.deep.equal([3, 2, 1]);
    });

    it('is empty when stop is not reachable in the step direction', () => {
      expect([...utils.range(3, 0)]).to.deep.equal([]);
      expect([...utils.range(0, 3, -1)]).to.deep.equal([]);
    });
  });
});
