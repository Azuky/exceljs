const testXformHelper = require('../test-xform-helper');

const ExtXform = verquire('xlsx/xform/drawing/ext-xform');

// 1 pixel at 96 dpi is 9525 EMU
const expectations = [
  {
    title: 'pixels to EMU',
    create() {
      return new ExtXform({ tag: 'xdr:ext' });
    },
    preparedModel: { width: 100, height: 200 },
    xml: '<xdr:ext cx="952500" cy="1905000"/>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse', 'parseIn'],
  },
  {
    title: 'fractional pixels are floored when written',
    create() {
      return new ExtXform({ tag: 'xdr:ext' });
    },
    preparedModel: { width: 10.5, height: 0.1 },
    xml: '<xdr:ext cx="100012" cy="952"/>',
    tests: ['render'],
  },
  {
    title: 'missing size attributes are zero',
    create() {
      return new ExtXform({ tag: 'xdr:ext' });
    },
    xml: '<xdr:ext/>',
    parsedModel: { width: 0, height: 0 },
    tests: ['parse'],
  },
];

describe('ExtXform (drawing)', () => {
  testXformHelper(expectations);
});
