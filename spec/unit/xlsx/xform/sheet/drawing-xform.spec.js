const testXformHelper = require('../test-xform-helper');

const DrawingXform = verquire('xlsx/xform/sheet/drawing-xform');

const expectations = [
  {
    title: 'Drawing relationship',
    create() {
      return new DrawingXform();
    },
    preparedModel: { rId: 'rId3' },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<drawing r:id="rId3"/>',
    tests: ['render', 'renderIn', 'parse', 'parseIn'],
  },
];

describe('DrawingXform', () => {
  testXformHelper(expectations);
});
