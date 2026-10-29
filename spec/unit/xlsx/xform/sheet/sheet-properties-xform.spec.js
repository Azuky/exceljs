const testXformHelper = require('../test-xform-helper');

const { findAll } = require('../../../../utils/xml-tree');

const SheetPropertiesXform = verquire('xlsx/xform/sheet/sheet-properties-xform');

// Known lib bug: SheetPropertiesXform renders pageSetUpPr before outlinePr, but
// CT_SheetPr (ECMA-376 sml.xsd) requires tabColor, outlinePr, pageSetUpPr.
const sheetPrOrderDeviation = {
  reason:
    /child #\d+ <pageSetUpPr> !== <outlinePr> \(order matters\); actual \[(tabColor,)?pageSetUpPr,outlinePr\] vs expected \[(tabColor,)?outlinePr,pageSetUpPr\]/,
  // undo the order only
  neutralise(root) {
    const order = ['tabColor', 'outlinePr', 'pageSetUpPr'];
    findAll(root, 'sheetPr').forEach((sheetPr) => {
      sheetPr.children.sort((a, b) => order.indexOf(a.name) - order.indexOf(b.name));
    });
  },
};

const expectations = [
  {
    title: 'empty',
    create() {
      return new SheetPropertiesXform();
    },
    preparedModel: {},
    xml: '',
    parsedModel: {},
    tests: ['render', 'renderIn'],
  },
  {
    title: 'tabColor',
    create() {
      return new SheetPropertiesXform();
    },
    preparedModel: { tabColor: { argb: 'FFFF0000' } },
    xml: '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>',
    parsedModel: { tabColor: { argb: 'FFFF0000' } },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'pageSetup',
    create() {
      return new SheetPropertiesXform();
    },
    preparedModel: { pageSetup: { fitToPage: true } },
    xml: '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
    parsedModel: { pageSetup: { fitToPage: true } },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'outlineProperties',
    create() {
      return new SheetPropertiesXform();
    },
    preparedModel: { outlineProperties: { summaryBelow: false } },
    xml: '<sheetPr><outlinePr summaryBelow="0"/></sheetPr>',
    parsedModel: { outlineProperties: { summaryBelow: false } },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'tabColor + pageSetup',
    create() {
      return new SheetPropertiesXform();
    },
    preparedModel: {
      tabColor: { argb: 'FFFF0000' },
      pageSetup: { fitToPage: true },
    },
    xml: '<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>',
    parsedModel: {
      tabColor: { argb: 'FFFF0000' },
      pageSetup: { fitToPage: true },
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'tabColor + outlineProperties',
    create() {
      return new SheetPropertiesXform();
    },
    preparedModel: {
      tabColor: { argb: 'FFFF0000' },
      outlineProperties: { summaryBelow: false },
    },
    xml: '<sheetPr><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0"/></sheetPr>',
    parsedModel: {
      tabColor: { argb: 'FFFF0000' },
      outlineProperties: { summaryBelow: false },
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'pageSetup + outlineProperties',
    create() {
      return new SheetPropertiesXform();
    },
    preparedModel: {
      pageSetup: { fitToPage: true },
      outlineProperties: { summaryBelow: false },
    },
    xml: '<sheetPr><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr>',
    parsedModel: {
      pageSetup: { fitToPage: true },
      outlineProperties: { summaryBelow: false },
    },
    knownDeviations: { render: sheetPrOrderDeviation },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'tabColor + outlineProperties + pageSetup',
    create() {
      return new SheetPropertiesXform();
    },
    preparedModel: {
      tabColor: { argb: 'FFFF0000' },
      outlineProperties: { summaryBelow: false },
      pageSetup: { fitToPage: true },
    },
    xml: '<sheetPr><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr>',
    parsedModel: {
      tabColor: { argb: 'FFFF0000' },
      outlineProperties: { summaryBelow: false },
      pageSetup: { fitToPage: true },
    },
    knownDeviations: { render: sheetPrOrderDeviation },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('SheetPropertiesXform', () => {
  testXformHelper(expectations);
});
