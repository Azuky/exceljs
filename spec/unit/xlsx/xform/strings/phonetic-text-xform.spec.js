const testXformHelper = require('../test-xform-helper');

const PhoneticTextXform = verquire('xlsx/xform/strings/phonetic-text-xform');

const expectations = [
  {
    title: 'text',
    create() {
      return new PhoneticTextXform();
    },
    preparedModel: { text: 'Hello, World!', sb: 0, eb: 1 },
    xml: '<rPh sb="0" eb="1"><t>Hello, World!</t></rPh>',
    parsedModel: { text: 'Hello, World!', sb: 0, eb: 1 },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Katakana',
    create() {
      return new PhoneticTextXform();
    },
    preparedModel: { sb: 0, eb: 2, text: 'ヤクワリ' },
    xml: '<rPh sb="0" eb="2"><t>ヤクワリ</t></rPh>',
    parsedModel: { sb: 0, eb: 2, text: 'ヤクワリ' },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'rich text runs',
    create() {
      return new PhoneticTextXform();
    },
    preparedModel: {
      sb: 0,
      eb: 2,
      richText: [{ text: 'カン' }, { font: { bold: true }, text: 'ジ' }],
    },
    xml: '<rPh sb="0" eb="2"><r><t>カン</t></r><r><rPr><b/></rPr><t>ジ</t></r></rPh>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse', 'parseIn'],
  },
  {
    title: 'missing offsets are written as 0',
    create() {
      return new PhoneticTextXform();
    },
    preparedModel: { text: 'ヤク' },
    xml: '<rPh sb="0" eb="0"><t>ヤク</t></rPh>',
    tests: ['render'],
  },
  {
    title: 'text with surrounding spaces',
    create() {
      return new PhoneticTextXform();
    },
    preparedModel: { sb: 0, eb: 1, text: ' ヤ ' },
    xml: '<rPh sb="0" eb="1"><t xml:space="preserve"> ヤ </t></rPh>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('PhoneticTextXform', () => {
  testXformHelper(expectations);
});
