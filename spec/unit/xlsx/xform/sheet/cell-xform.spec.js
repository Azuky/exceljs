const testXformHelper = require('../test-xform-helper');

const CellXform = verquire('xlsx/xform/sheet/cell-xform');
const SharedStringsXform = verquire('xlsx/xform/strings/shared-strings-xform');
const Enums = verquire('doc/enums');

const fakeStyles = {
  addStyleModel(style, effectiveType) {
    if (effectiveType === Enums.ValueType.Date) {
      return 1;
    }
    return 0;
  },
  getStyleModel(styleId) {
    switch (styleId) {
      case 1:
        return { numFmt: 'mm-dd-yy' };
      default:
        return null;
    }
  },
};

const fakeHyperlinkMap = {
  H1: 'http://www.foo.com',
};

const formulaOptions = () => ({
  sharedStrings: new SharedStringsXform(),
  hyperlinks: [],
  hyperlinkMap: fakeHyperlinkMap,
  styles: fakeStyles,
  formulae: {},
  siFormulae: 0,
});

const expectations = [
  {
    title: 'Styled Null',
    create() {
      return new CellXform();
    },
    preparedModel: { address: 'A1', type: Enums.ValueType.Null, styleId: 1 },
    xml: '<c r="A1" s="1" />',
    parsedModel: { address: 'A1', type: Enums.ValueType.Null, styleId: 1 },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Number',
    create() {
      return new CellXform();
    },
    preparedModel: { address: 'A1', type: Enums.ValueType.Number, value: 5 },
    parsedModel: { address: 'A1', type: Enums.ValueType.Number, value: 5 },
    xml: '<c r="A1"><v>5</v></c>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Boolean',
    create() {
      return new CellXform();
    },
    preparedModel: { address: 'A1', type: Enums.ValueType.Boolean, value: true },
    parsedModel: { address: 'A1', type: Enums.ValueType.Boolean, value: true },
    xml: '<c r="A1" t="b"><v>1</v></c>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Error',
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.Error,
      value: { error: '#N/A' },
    },
    parsedModel: {
      address: 'A1',
      type: Enums.ValueType.Error,
      value: { error: '#N/A' },
    },
    xml: '<c r="A1" t="e"><v>#N/A</v></c>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'String',
    create() {
      return new CellXform();
    },
    initialModel: { address: 'A1', type: Enums.ValueType.String, value: 'Foo' },
    preparedModel: { address: 'A1', type: Enums.ValueType.String, value: 'Foo' },
    xml: '<c r="A1" t="str"><v>Foo</v></c>',
    parsedModel: { address: 'A1', type: Enums.ValueType.String, value: 'Foo' },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: 'Foo',
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles },
  },
  {
    title: 'String with Invalid Number',
    create() {
      return new CellXform();
    },
    initialModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: '6E1000',
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: '6E1000',
    },
    xml: '<c r="A1" t="str"><v>6E1000</v></c>',
    parsedModel: { address: 'A1', type: Enums.ValueType.String, value: '6E1000' },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: '6E1000',
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: {
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
    },
  },
  {
    title: 'Inline String with plain text',
    create() {
      return new CellXform();
    },
    xml: '<c r="A1" t="inlineStr"><is><t>Foo</t></is></c>',
    parsedModel: { address: 'A1', type: Enums.ValueType.String, value: 'Foo' },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: 'Foo',
    },
    tests: ['parse', 'reconcile'],
    options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles },
  },
  {
    title: 'Inline String with RichText',
    create() {
      return new CellXform();
    },
    initialModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: {
        richText: [
          { font: { color: { argb: 'FF0000' } }, text: 'red' },
          { font: { color: { argb: '00FF00' } }, text: 'green' },
        ],
      },
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: {
        richText: [
          { font: { color: { argb: 'FF0000' } }, text: 'red' },
          { font: { color: { argb: '00FF00' } }, text: 'green' },
        ],
      },
    },
    xml: '<c r="A1" t="inlineStr"><is><r><rPr><color rgb="FF0000"/></rPr><t>red</t></r><r><rPr><color rgb="00FF00"/></rPr><t>green</t></r></is></c>',
    parsedModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: {
        richText: [
          { font: { color: { argb: 'FF0000' } }, text: 'red' },
          { font: { color: { argb: '00FF00' } }, text: 'green' },
        ],
      },
    },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.RichText,
      value: {
        richText: [
          { font: { color: { argb: 'FF0000' } }, text: 'red' },
          { font: { color: { argb: '00FF00' } }, text: 'green' },
        ],
      },
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles },
  },
  {
    title: 'Shared String',
    create() {
      return new CellXform();
    },
    initialModel: { address: 'A1', type: Enums.ValueType.String, value: 'Foo' },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: 'Foo',
      ssId: 0,
    },
    xml: '<c r="A1" t="s"><v>0</v></c>',
    parsedModel: { address: 'A1', type: Enums.ValueType.String, value: 0 },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: 'Foo',
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
    },
  },
  {
    title: 'Shared String with RichText',
    create() {
      return new CellXform();
    },
    initialModel: {
      address: 'A1',
      type: Enums.ValueType.RichText,
      value: {
        richText: [
          { font: { color: { argb: 'FF0000' } }, text: 'red' },
          { font: { color: { argb: '00FF00' } }, text: 'green' },
        ],
      },
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.RichText,
      value: {
        richText: [
          { font: { color: { argb: 'FF0000' } }, text: 'red' },
          { font: { color: { argb: '00FF00' } }, text: 'green' },
        ],
      },
      ssId: 0,
    },
    xml: '<c r="A1" t="s"><v>0</v></c>',
    parsedModel: {
      address: 'A1',
      type: Enums.ValueType.String,
      value: 0,
    },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.RichText,
      value: {
        richText: [
          { font: { color: { argb: 'FF0000' } }, text: 'red' },
          { font: { color: { argb: '00FF00' } }, text: 'green' },
        ],
      },
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
    },
  },
  {
    title: 'Date',
    create() {
      return new CellXform();
    },
    initialModel: {
      address: 'A1',
      type: Enums.ValueType.Date,
      value: new Date('2016-06-09T00:00:00.000Z'),
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.Date,
      value: new Date('2016-06-09T00:00:00.000Z'),
      styleId: 1,
    },
    xml: '<c r="A1" s="1"><v>42530</v></c>',
    parsedModel: {
      address: 'A1',
      type: Enums.ValueType.Number,
      value: 42530,
      styleId: 1,
    },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.Date,
      value: new Date('2016-06-09T00:00:00.000Z'),
      style: { numFmt: 'mm-dd-yy' },
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
    },
  },
  {
    title: 'Hyperlink',
    create() {
      return new CellXform();
    },
    initialModel: {
      address: 'H1',
      type: Enums.ValueType.Hyperlink,
      hyperlink: 'http://www.foo.com',
      text: 'www.foo.com',
    },
    preparedModel: {
      address: 'H1',
      type: Enums.ValueType.Hyperlink,
      hyperlink: 'http://www.foo.com',
      text: 'www.foo.com',
      ssId: 0,
    },
    xml: '<c r="H1" t="s"><v>0</v></c>',
    parsedModel: { address: 'H1', type: Enums.ValueType.String, value: 0 },
    reconciledModel: {
      address: 'H1',
      type: Enums.ValueType.Hyperlink,
      text: 'www.foo.com',
      hyperlink: 'http://www.foo.com',
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
    },
  },
  {
    title: 'String Formula',
    create() {
      return new CellXform();
    },
    initialModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'A2',
      result: 'Foo',
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'A2',
      result: 'Foo',
    },
    xml: '<c r="A1" t="str"><f>A2</f><v>Foo</v></c>',
    parsedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'A2',
      result: 'Foo',
    },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'A2',
      result: 'Foo',
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Number Formula',
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'A2',
      result: 7,
    },
    xml: '<c r="A1"><f>A2</f><v>7</v></c>',
    parsedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'A2',
      result: 7,
    },
    tests: ['render', 'renderIn', 'parse'],
    options: { formulae: {}, siFormulae: 0 },
  },
  {
    title: 'Master Shared Formula',
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      shareType: 'shared',
      ref: 'A2:B2',
      formula: 'A1',
      result: 2,
      si: 0,
    },
    xml: '<c r="A2"><f t="shared" ref="A2:B2" si="0">A1</f><v>2</v></c>',
    parsedModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      shareType: 'shared',
      ref: 'A2:B2',
      formula: 'A1',
      result: 2,
      si: '0',
    },
    reconciledModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      shareType: 'shared',
      ref: 'A2:B2',
      formula: 'A1',
      result: 2,
    },
    tests: ['render', 'renderIn', 'parse', 'reconcile'],
    options: {
      styles: fakeStyles,
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Shared Formula Slave',
    create() {
      return new CellXform();
    },
    initialModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      sharedFormula: 'A1',
      result: 2,
    },
    preparedModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      sharedFormula: 'A1',
      result: 2,
      si: 0,
    },
    xml: '<c r="A2"><f t="shared" si="0" /><v>2</v></c>',
    parsedModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      result: 2,
      shareType: 'shared',
      si: '0',
    },
    reconciledModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      result: 2,
      sharedFormula: 'A1',
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: {
      styles: fakeStyles,
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      formulae: {
        A1: {
          address: 'A1',
          type: Enums.ValueType.Formula,
          formula: 'ROW()',
          result: 1,
          si: 0,
        },
        0: 'A1',
      },
      siFormulae: 1,
    },
  },
  {
    title: 'Array Shared Formula',
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      shareType: 'array',
      ref: 'A2:B2',
      formula: 'A1',
      result: 2,
    },
    xml: '<c r="A2"><f t="array" ref="A2:B2">A1</f><v>2</v></c>',
    parsedModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      shareType: 'array',
      ref: 'A2:B2',
      formula: 'A1',
      result: 2,
    },
    reconciledModel: {
      address: 'A2',
      type: Enums.ValueType.Formula,
      shareType: 'array',
      ref: 'A2:B2',
      formula: 'A1',
      result: 2,
    },
    tests: ['render', 'renderIn', 'parse', 'reconcile'],
    options: {
      styles: fakeStyles,
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Boolean Formula (true)',
    create() {
      return new CellXform();
    },
    initialModel: { address: 'A1', type: Enums.ValueType.Formula, formula: 'A1>1', result: true },
    preparedModel: { address: 'A1', type: Enums.ValueType.Formula, formula: 'A1>1', result: true },
    xml: '<c r="A1" t="b"><f>A1&gt;1</f><v>1</v></c>',
    parsedModel: { address: 'A1', type: Enums.ValueType.Formula, formula: 'A1>1', result: true },
    reconciledModel: { address: 'A1', type: Enums.ValueType.Formula, formula: 'A1>1', result: true },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: formulaOptions(),
  },
  {
    title: 'Boolean Formula (false)',
    create() {
      return new CellXform();
    },
    preparedModel: { address: 'A1', type: Enums.ValueType.Formula, formula: 'A1>1', result: false },
    xml: '<c r="A1" t="b"><f>A1&gt;1</f><v>0</v></c>',
    parsedModel: { address: 'A1', type: Enums.ValueType.Formula, formula: 'A1>1', result: false },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Error Formula',
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: '1/0',
      result: { error: '#DIV/0!' },
    },
    xml: '<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>',
    parsedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: '1/0',
      result: { error: '#DIV/0!' },
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Formula without a result',
    create() {
      return new CellXform();
    },
    preparedModel: { address: 'A1', type: Enums.ValueType.Formula, formula: 'SUM(B1:B2)' },
    xml: '<c r="A1"><f>SUM(B1:B2)</f></c>',
    parsedModel: { address: 'A1', type: Enums.ValueType.Formula, formula: 'SUM(B1:B2)' },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Date Formula',
    create() {
      return new CellXform();
    },
    initialModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'B1+1',
      result: new Date('2016-06-09T00:00:00.000Z'),
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'B1+1',
      result: new Date('2016-06-09T00:00:00.000Z'),
      styleId: 1,
    },
    xml: '<c r="A1" s="1"><f>B1+1</f><v>42530</v></c>',
    parsedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'B1+1',
      result: 42530,
      styleId: 1,
    },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'B1+1',
      result: new Date('2016-06-09T00:00:00.000Z'),
      style: { numFmt: 'mm-dd-yy' },
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: formulaOptions(),
  },
  {
    // in the 1904 date system serial 42530 is 1462 days later than in the 1900 system
    title: 'Date Formula in the 1904 date system',
    create() {
      return new CellXform();
    },
    initialModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'B1+1',
      result: new Date('2020-06-10T00:00:00.000Z'),
    },
    preparedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'B1+1',
      result: new Date('2020-06-10T00:00:00.000Z'),
      styleId: 1,
      date1904: true,
    },
    xml: '<c r="A1" s="1"><f>B1+1</f><v>42530</v></c>',
    parsedModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'B1+1',
      result: 42530,
      styleId: 1,
    },
    reconciledModel: {
      address: 'A1',
      type: Enums.ValueType.Formula,
      formula: 'B1+1',
      result: new Date('2020-06-10T00:00:00.000Z'),
      style: { numFmt: 'mm-dd-yy' },
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options: { ...formulaOptions(), date1904: true },
  },
];

describe('CellXform', () => {
  testXformHelper(expectations);

  describe('error paths', () => {
    it('throws when a shared formula clone has no master', () => {
      const xform = new CellXform();
      const model = { address: 'B2', type: Enums.ValueType.Formula, sharedFormula: 'Z9', result: 1 };
      expect(() => xform.prepare(model, formulaOptions())).to.throw(
        'Shared Formula master must exist above and or left of clone for cell B2',
      );
    });

    it('throws when a formula result has an unsupported type', () => {
      const xform = new CellXform();
      const model = { address: 'A1', type: Enums.ValueType.Formula, formula: 'A2', result: { unexpected: true } };
      expect(() => xform.prepare(model, formulaOptions())).to.throw('I could not understand type of value');
    });

    it('joins value text that arrives in several text events', () => {
      const xform = new CellXform();
      xform.parseOpen({ name: 'c', attributes: { r: 'A1', t: 'str' } });
      xform.parseOpen({ name: 'v', attributes: {} });
      xform.parseText('Hello, ');
      xform.parseText('world');
      xform.parseClose('v');
      xform.parseClose('c');
      expect(xform.model).to.deep.include({ address: 'A1', type: Enums.ValueType.String, value: 'Hello, world' });
    });

    it('joins formula text that arrives in several text events', () => {
      const xform = new CellXform();
      xform.parseOpen({ name: 'c', attributes: { r: 'A1' } });
      xform.parseOpen({ name: 'f', attributes: {} });
      xform.parseText('SUM(A2:');
      xform.parseText('A9)');
      xform.parseClose('f');
      xform.parseClose('c');
      expect(xform.model.formula).to.equal('SUM(A2:A9)');
      expect(xform.model.type).to.equal(Enums.ValueType.Formula);
    });
  });
});
