const testXformHelper = require('../test-xform-helper');

const DefinedNameXform = verquire('xlsx/xform/book/defined-name-xform');

const expectations = [
  {
    title: 'Defined Names',
    create() {
      return new DefinedNameXform();
    },
    preparedModel: { name: 'foo', ranges: ['bar!$A$1:$C$1'] },
    xml: '<definedName name="foo">bar!$A$1:$C$1</definedName>',
    parsedModel: { name: 'foo', ranges: ['bar!$A$1:$C$1'] },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Print Area',
    create() {
      return new DefinedNameXform();
    },
    preparedModel: {
      name: '_xlnm.Print_Area',
      localSheetId: 0,
      ranges: ['bar!$A$1:$C$10'],
    },
    xml: '<definedName name="_xlnm.Print_Area" localSheetId="0">bar!$A$1:$C$10</definedName>',
    parsedModel: {
      name: '_xlnm.Print_Area',
      localSheetId: 0,
      ranges: ['bar!$A$1:$C$10'],
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'String with something that looks like a range',
    create() {
      return new DefinedNameXform();
    },
    preparedModel: { name: 'foo', ranges: [] },
    xml: '<definedName name="foo">"OFFSET($A$10;0;0;0;1)"</definedName>',
    parsedModel: { name: 'foo', ranges: [] },
    tests: ['parse'],
  },
  {
    title: 'Quoted sheet name containing a comma',
    create() {
      return new DefinedNameXform();
    },
    xml: `<definedName name="n">'My, Sheet'!$A$1:$B$2</definedName>`,
    parsedModel: { name: 'n', ranges: ["'My, Sheet'!$A$1:$B$2"] },
    tests: ['parse'],
  },
  {
    title: 'Several ranges, one with a quoted comma',
    create() {
      return new DefinedNameXform();
    },
    xml: `<definedName name="n">'My, Sheet'!$A$1:$B$2,Sheet2!$C$3</definedName>`,
    parsedModel: { name: 'n', ranges: ["'My, Sheet'!$A$1:$B$2", 'Sheet2!$C$3'] },
    tests: ['parse'],
  },
  {
    title: 'Quoted sheet name without a comma',
    create() {
      return new DefinedNameXform();
    },
    xml: `<definedName name="n">'My Sheet'!$A$1,Sheet2!$C$3</definedName>`,
    parsedModel: { name: 'n', ranges: ["'My Sheet'!$A$1", 'Sheet2!$C$3'] },
    tests: ['parse'],
  },
  {
    title: 'Quoted sheet name with two commas',
    create() {
      return new DefinedNameXform();
    },
    xml: `<definedName name="n">'a,b,c'!$A$1</definedName>`,
    parsedModel: { name: 'n', ranges: ["'a,b,c'!$A$1"] },
    tests: ['parse'],
  },
  {
    title: 'Unbalanced quote is dropped',
    create() {
      return new DefinedNameXform();
    },
    xml: `<definedName name="n">'Unclosed!$A$1</definedName>`,
    parsedModel: { name: 'n', ranges: [] },
    tests: ['parse'],
  },
  {
    title: 'Invalid range among valid ones is dropped',
    create() {
      return new DefinedNameXform();
    },
    xml: `<definedName name="n">Sheet1!$A$1,NOT A RANGE,Sheet2!$B$2</definedName>`,
    parsedModel: { name: 'n', ranges: ['Sheet1!$A$1', 'Sheet2!$B$2'] },
    tests: ['parse'],
  },
  {
    title: 'Sheet scoped name',
    create() {
      return new DefinedNameXform();
    },
    xml: `<definedName name="n" localSheetId="2">Sheet3!$A$1</definedName>`,
    parsedModel: { name: 'n', localSheetId: 2, ranges: ['Sheet3!$A$1'] },
    tests: ['parse'],
  },
];

describe('DefinedNameXform', () => {
  testXformHelper(expectations);
});
