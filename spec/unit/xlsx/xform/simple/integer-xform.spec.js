const testXformHelper = require('../test-xform-helper');

const IntegerXform = verquire('xlsx/xform/simple/integer-xform');

const expectations = [
  {
    title: 'five',
    create() {
      return new IntegerXform({ tag: 'integer', attr: 'val' });
    },
    preparedModel: 5,
    xml: '<integer val="5"/>',
    parsedModel: 5,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'zero',
    create() {
      return new IntegerXform({ tag: 'integer', attr: 'val' });
    },
    preparedModel: 0,
    xml: '',
    parsedModel: undefined,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'undefined',
    create() {
      return new IntegerXform({ tag: 'integer', attr: 'val' });
    },
    preparedModel: undefined,
    xml: '',
    parsedModel: undefined,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'text content',
    create() {
      return new IntegerXform({ tag: 'i' });
    },
    preparedModel: 7,
    xml: '<i>7</i>',
    parsedModel: 7,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'zero rendered when requested',
    create() {
      return new IntegerXform({ tag: 'integer', attr: 'val', zero: true });
    },
    preparedModel: 0,
    xml: '<integer val="0"/>',
    parsedModel: 0,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'empty text content is zero',
    create() {
      return new IntegerXform({ tag: 'i' });
    },
    xml: '<i></i>',
    parsedModel: 0,
    tests: ['parse'],
  },
  {
    title: 'extra attributes next to the value',
    create() {
      return new IntegerXform({ tag: 'integer', attr: 'val', attrs: { unit: 'pt' } });
    },
    preparedModel: 3,
    xml: '<integer unit="pt" val="3"/>',
    parsedModel: 3,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'not a number in the attribute',
    create() {
      return new IntegerXform({ tag: 'integer', attr: 'val' });
    },
    xml: '<integer val="abc"/>',
    parsedModel: NaN,
    tests: ['parse'],
  },
];

describe('IntegerXform', () => {
  testXformHelper(expectations);
});
