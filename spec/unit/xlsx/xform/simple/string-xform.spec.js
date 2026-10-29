const testXformHelper = require('../test-xform-helper');

const StringXform = verquire('xlsx/xform/simple/string-xform');

const expectations = [
  {
    title: 'hello',
    create() {
      return new StringXform({ tag: 'string', attr: 'val' });
    },
    preparedModel: 'Hello, World!',
    xml: '<string val="Hello, World!"/>',
    parsedModel: 'Hello, World!',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'empty',
    create() {
      return new StringXform({ tag: 'string', attr: 'val' });
    },
    preparedModel: '',
    xml: '<string val=""/>',
    parsedModel: '',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'undefined',
    create() {
      return new StringXform({ tag: 'string', attr: 'val' });
    },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'text content',
    create() {
      return new StringXform({ tag: 's' });
    },
    preparedModel: 'Hello, World!',
    xml: '<s>Hello, World!</s>',
    parsedModel: 'Hello, World!',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'text content with markup characters',
    create() {
      return new StringXform({ tag: 's' });
    },
    preparedModel: 'a & b < c',
    xml: '<s>a &amp; b &lt; c</s>',
    parsedModel: 'a & b < c',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'empty text content',
    create() {
      return new StringXform({ tag: 's' });
    },
    preparedModel: '',
    xml: '<s></s>',
    parsedModel: '',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'extra attributes next to the value',
    create() {
      return new StringXform({ tag: 'string', attr: 'val', attrs: { lang: 'en' } });
    },
    preparedModel: 'Hello',
    xml: '<string lang="en" val="Hello"/>',
    parsedModel: 'Hello',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'extra attributes with text content',
    create() {
      return new StringXform({ tag: 's', attrs: { lang: 'en' } });
    },
    preparedModel: 'Hello',
    xml: '<s lang="en">Hello</s>',
    parsedModel: 'Hello',
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('StringXform', () => {
  testXformHelper(expectations);
});
