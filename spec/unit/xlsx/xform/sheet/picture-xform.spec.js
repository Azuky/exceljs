const testXformHelper = require('../test-xform-helper');

const PictureXform = verquire('xlsx/xform/sheet/picture-xform');

const expectations = [
  {
    title: 'Picture relationship',
    create() {
      return new PictureXform();
    },
    preparedModel: { rId: 'rId3' },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<picture r:id="rId3"/>',
    tests: ['render', 'renderIn', 'parse', 'parseIn'],
  },
];

describe('PictureXform', () => {
  testXformHelper(expectations);
});
