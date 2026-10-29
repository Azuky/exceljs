const testXformHelper = require('../test-xform-helper');

const SheetXform = verquire('xlsx/xform/book/sheet-xform');

const expectations = [
  {
    title: 'Sheet entry',
    create() {
      return new SheetXform();
    },
    preparedModel: { id: 2, name: 'Data', rId: 'rId2' },
    get parsedModel() {
      return { ...this.preparedModel };
    },
    xml: '<sheet sheetId="2" name="Data" r:id="rId2"/>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Hidden sheet with xml-encoded name',
    create() {
      return new SheetXform();
    },
    preparedModel: { id: 5, name: 'R&D <2>', state: 'hidden', rId: 'rId7' },
    get parsedModel() {
      return { ...this.preparedModel };
    },
    xml: '<sheet sheetId="5" name="R&amp;D &lt;2&gt;" state="hidden" r:id="rId7"/>',
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('SheetXform (workbook)', () => {
  testXformHelper(expectations);
});
