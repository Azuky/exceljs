const testXformHelper = require('../test-xform-helper');

const FloatXform = verquire('xlsx/xform/simple/float-xform');

const expectations = [
  {
    title: 'five',
    create() {
      return new FloatXform({ tag: 'float', attr: 'val' });
    },
    preparedModel: 5,
    xml: '<float val="5"/>',
    parsedModel: 5,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'pi',
    create() {
      return new FloatXform({ tag: 'float', attr: 'val' });
    },
    preparedModel: 3.14,
    xml: '<float val="3.14"/>',
    parsedModel: 3.14,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'zero',
    create() {
      return new FloatXform({ tag: 'float', attr: 'val' });
    },
    preparedModel: 0,
    xml: '<float val="0"/>',
    parsedModel: 0,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'undefined',
    create() {
      return new FloatXform({ tag: 'float', attr: 'val' });
    },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'text content',
    create() {
      return new FloatXform({ tag: 'f' });
    },
    preparedModel: 1.5,
    xml: '<f>1.5</f>',
    parsedModel: 1.5,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'text content zero',
    create() {
      return new FloatXform({ tag: 'f' });
    },
    preparedModel: 0,
    xml: '<f>0</f>',
    parsedModel: 0,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'extra attributes next to the value',
    create() {
      return new FloatXform({ tag: 'float', attr: 'val', attrs: { unit: 'pt' } });
    },
    preparedModel: 1.5,
    xml: '<float unit="pt" val="1.5"/>',
    parsedModel: 1.5,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'extra attributes with text content',
    create() {
      return new FloatXform({ tag: 'f', attrs: { unit: 'pt' } });
    },
    preparedModel: 2.25,
    xml: '<f unit="pt">2.25</f>',
    parsedModel: 2.25,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'not a number in the attribute',
    create() {
      return new FloatXform({ tag: 'float', attr: 'val' });
    },
    xml: '<float val="abc"/>',
    parsedModel: NaN,
    tests: ['parse'],
  },
];

describe('FloatXform', () => {
  testXformHelper(expectations);
});
