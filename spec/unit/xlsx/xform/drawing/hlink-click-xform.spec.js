const testXformHelper = require('../test-xform-helper');

const HLinkClickXform = verquire('xlsx/xform/drawing/hlink-click-xform');

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const expectations = [
  {
    title: 'hyperlink with tooltip',
    create() {
      return new HLinkClickXform();
    },
    preparedModel: { hyperlinks: { rId: 'rId3', tooltip: 'exceljs' } },
    xml: `<a:hlinkClick xmlns:r="${R_NS}" r:id="rId3" tooltip="exceljs"/>`,
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse', 'parseIn'],
  },
  {
    title: 'hyperlink without tooltip',
    create() {
      return new HLinkClickXform();
    },
    preparedModel: { hyperlinks: { rId: 'rId4' } },
    xml: `<a:hlinkClick xmlns:r="${R_NS}" r:id="rId4"/>`,
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'picture without a hyperlink renders nothing',
    create() {
      return new HLinkClickXform();
    },
    preparedModel: {},
    xml: '',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'hyperlink without a relationship id renders nothing',
    create() {
      return new HLinkClickXform();
    },
    preparedModel: { hyperlinks: { tooltip: 'orphan' } },
    xml: '',
    tests: ['render', 'renderIn'],
  },
];

describe('HLinkClickXform', () => {
  testXformHelper(expectations);
});
