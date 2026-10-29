const testXformHelper = require('../test-xform-helper');

const TablePartXform = verquire('xlsx/xform/sheet/table-part-xform');

const expectations = [
  {
    title: 'TablePart relationship',
    create() {
      return new TablePartXform();
    },
    preparedModel: { rId: 'rId3' },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<tablePart r:id="rId3"/>',
    tests: ['render', 'renderIn', 'parse', 'parseIn'],
  },
];

describe('TablePartXform', () => {
  testXformHelper(expectations);
});
