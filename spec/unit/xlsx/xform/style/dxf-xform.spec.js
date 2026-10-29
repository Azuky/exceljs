const testXformHelper = require('../test-xform-helper');

const DxfXform = verquire('xlsx/xform/style/dxf-xform');

// the parser reports the parts that are absent from the <dxf> as null
const parsed = (parts) => ({
  alignment: null,
  border: null,
  fill: null,
  font: null,
  numFmt: null,
  protection: null,
  ...parts,
});

const expectations = [
  {
    title: 'Font only',
    create: () => new DxfXform(),
    preparedModel: { font: { bold: true, color: { argb: 'FFFF0000' } } },
    xml: '<dxf><font><b/><color rgb="FFFF0000"/></font></dxf>',
    parsedModel: parsed({ font: { bold: true, color: { argb: 'FFFF0000' } } }),
    tests: ['render', 'renderIn', 'parse', 'parseIn'],
  },
  {
    // child order follows CT_Dxf: font, numFmt, fill, alignment, border, protection
    title: 'All parts',
    create: () => new DxfXform(),
    preparedModel: {
      font: { bold: true },
      numFmt: '0.00%',
      numFmtId: 165,
      fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFF00' } },
      alignment: { horizontal: 'center' },
      border: { left: { style: 'thin' } },
      protection: { locked: false },
    },
    xml: `<dxf>
      <font><b/></font>
      <numFmt numFmtId="165" formatCode="0.00%"/>
      <fill><patternFill patternType="solid"><bgColor rgb="FFFFFF00"/></patternFill></fill>
      <alignment horizontal="center"/>
      <border><left style="thin"/><right/><top/><bottom/><diagonal/></border>
      <protection locked="0"/>
    </dxf>`,
    parsedModel: parsed({
      font: { bold: true },
      numFmt: { id: 165, formatCode: '0.00%' },
      fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFF00' } },
      alignment: { horizontal: 'center' },
      border: { left: { style: 'thin' } },
      protection: { locked: false, hidden: false },
    }),
    tests: ['render', 'renderIn', 'parse', 'parseIn'],
  },
  {
    title: 'Alignment only',
    create: () => new DxfXform(),
    preparedModel: { alignment: { horizontal: 'right', vertical: 'top', wrapText: true } },
    xml: '<dxf><alignment horizontal="right" vertical="top" wrapText="1"/></dxf>',
    get parsedModel() {
      return parsed(this.preparedModel);
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Border only',
    create: () => new DxfXform(),
    preparedModel: {
      border: {
        left: { style: 'thin' },
        right: { style: 'thin' },
        top: { style: 'thin' },
        bottom: { style: 'double', color: { argb: 'FF00FF00' } },
      },
    },
    xml: `<dxf><border>
      <left style="thin"/><right style="thin"/><top style="thin"/>
      <bottom style="double"><color rgb="FF00FF00"/></bottom><diagonal/>
    </border></dxf>`,
    get parsedModel() {
      return parsed(this.preparedModel);
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Protection only',
    create: () => new DxfXform(),
    preparedModel: { protection: { locked: false, hidden: true } },
    xml: '<dxf><protection locked="0" hidden="1"/></dxf>',
    get parsedModel() {
      return parsed(this.preparedModel);
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Number format without an id is not rendered',
    create: () => new DxfXform(),
    preparedModel: { numFmt: '0.00%', font: { italic: true } },
    xml: '<dxf><font><i/></font></dxf>',
    tests: ['render'],
  },
];

describe('DxfXform', () => {
  testXformHelper(expectations);
});
