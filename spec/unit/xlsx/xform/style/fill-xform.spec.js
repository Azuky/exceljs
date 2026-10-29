const testXformHelper = require('../test-xform-helper');

const FillXform = verquire('xlsx/xform/style/fill-xform');

const expectations = [
  {
    title: 'Empty',
    create() {
      return new FillXform();
    },
    preparedModel: {},
    xml: '',
    parsedModel: undefined,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'None',
    create() {
      return new FillXform();
    },
    preparedModel: { type: 'pattern', pattern: 'none' },
    xml: '<fill><patternFill patternType="none"/></fill>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Gray 125',
    create() {
      return new FillXform();
    },
    preparedModel: { type: 'pattern', pattern: 'gray125' },
    xml: '<fill><patternFill patternType="gray125"/></fill>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Red Dark Vertical Pattern',
    create() {
      return new FillXform();
    },
    preparedModel: {
      type: 'pattern',
      pattern: 'darkVertical',
      fgColor: { argb: 'FFFF0000' },
    },
    xml: '<fill><patternFill patternType="darkVertical"><fgColor rgb="FFFF0000"/></patternFill></fill>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Red Green Dark Trellis Pattern',
    create() {
      return new FillXform();
    },
    preparedModel: {
      type: 'pattern',
      pattern: 'darkTrellis',
      fgColor: { argb: 'FFFF0000' },
      bgColor: { argb: 'FF00FF00' },
    },
    xml: '<fill><patternFill patternType="darkTrellis"><fgColor rgb="FFFF0000"/><bgColor rgb="FF00FF00"/></patternFill></fill>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Blue White Horizontal Gradient',
    create() {
      return new FillXform();
    },
    preparedModel: {
      type: 'gradient',
      gradient: 'angle',
      degree: 0,
      stops: [
        { position: 0, color: { argb: 'FF0000FF' } },
        { position: 1, color: { argb: 'FFFFFFFF' } },
      ],
    },
    xml:
      '<fill><gradientFill degree="0">' +
      '<stop position="0"><color rgb="FF0000FF"/></stop>' +
      '<stop position="1"><color rgb="FFFFFFFF"/></stop>' +
      '</gradientFill></fill>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'RGB Path Gradient',
    create() {
      return new FillXform();
    },
    preparedModel: {
      type: 'gradient',
      gradient: 'path',
      center: { left: 0.5, top: 0.5 },
      stops: [
        { position: 0, color: { argb: 'FFFF0000' } },
        { position: 0.5, color: { argb: 'FF00FF00' } },
        { position: 1, color: { argb: 'FF0000FF' } },
      ],
    },
    xml:
      '<fill><gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">' +
      '<stop position="0"><color rgb="FFFF0000"/></stop>' +
      '<stop position="0.5"><color rgb="FF00FF00"/></stop>' +
      '<stop position="1"><color rgb="FF0000FF"/></stop>' +
      '</gradientFill></fill>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Path Gradient with distinct right and bottom',
    create() {
      return new FillXform();
    },
    preparedModel: {
      type: 'gradient',
      gradient: 'path',
      center: { left: 0.5, top: 0.5, right: 1, bottom: 0.2 },
      stops: [
        { position: 0, color: { argb: 'FFFF0000' } },
        { position: 1, color: { argb: 'FF0000FF' } },
      ],
    },
    xml:
      '<fill><gradientFill type="path" left="0.5" right="1" top="0.5" bottom="0.2">' +
      '<stop position="0"><color rgb="FFFF0000"/></stop>' +
      '<stop position="1"><color rgb="FF0000FF"/></stop>' +
      '</gradientFill></fill>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Path Gradient with right and bottom left out mirrors left and top',
    create() {
      return new FillXform();
    },
    preparedModel: {
      type: 'gradient',
      gradient: 'path',
      center: { left: 0.25, top: 0.75 },
      stops: [
        { position: 0, color: { argb: 'FFFF0000' } },
        { position: 1, color: { argb: 'FF0000FF' } },
      ],
    },
    xml:
      '<fill><gradientFill type="path" left="0.25" right="0.25" top="0.75" bottom="0.75">' +
      '<stop position="0"><color rgb="FFFF0000"/></stop>' +
      '<stop position="1"><color rgb="FF0000FF"/></stop>' +
      '</gradientFill></fill>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    // attributes that are not written default to 0 (CT_GradientFill)
    title: 'Path Gradient from the top left corner',
    create() {
      return new FillXform();
    },
    preparedModel: {
      type: 'gradient',
      gradient: 'path',
      center: { left: 0, top: 0, right: 1, bottom: 1 },
      stops: [
        { position: 0, color: { argb: 'FFFF0000' } },
        { position: 1, color: { argb: 'FF0000FF' } },
      ],
    },
    xml:
      '<fill><gradientFill type="path" right="1" bottom="1">' +
      '<stop position="0"><color rgb="FFFF0000"/></stop>' +
      '<stop position="1"><color rgb="FF0000FF"/></stop>' +
      '</gradientFill></fill>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Path Gradient without right and bottom attributes parses them as 0',
    create() {
      return new FillXform();
    },
    xml:
      '<fill><gradientFill type="path" left="0.5" top="0.5">' +
      '<stop position="0"><color rgb="FFFF0000"/></stop>' +
      '<stop position="1"><color rgb="FF0000FF"/></stop>' +
      '</gradientFill></fill>',
    parsedModel: {
      type: 'gradient',
      gradient: 'path',
      center: { left: 0.5, top: 0.5, right: 0, bottom: 0 },
      stops: [
        { position: 0, color: { argb: 'FFFF0000' } },
        { position: 1, color: { argb: 'FF0000FF' } },
      ],
    },
    tests: ['parse'],
  },
];

describe('FillXform', () => {
  testXformHelper(expectations);
});
