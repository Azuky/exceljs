const testXformHelper = require('../test-xform-helper');

const CNvPrXform = verquire('xlsx/xform/drawing/c-nv-pr-xform');

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CREATION_ID =
  '<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">' +
  '<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{00000000-0008-0000-0000-000002000000}"/>' +
  '</a:ext></a:extLst>';

const expectations = [
  {
    title: 'picture with a hyperlink',
    create() {
      return new CNvPrXform();
    },
    preparedModel: { index: 1, hyperlinks: { rId: 'rId1', tooltip: 'Go' } },
    xml: `<xdr:cNvPr id="1" name="Picture 1"><a:hlinkClick xmlns:r="${R_NS}" r:id="rId1" tooltip="Go"/>${CREATION_ID}</xdr:cNvPr>`,
    parsedModel: { hyperlinks: { rId: 'rId1', tooltip: 'Go' } },
    tests: ['render', 'renderIn', 'parse', 'parseIn'],
  },
  {
    title: 'picture without a hyperlink',
    create() {
      return new CNvPrXform();
    },
    preparedModel: { index: 2 },
    xml: `<xdr:cNvPr id="2" name="Picture 2">${CREATION_ID}</xdr:cNvPr>`,
    tests: ['render', 'renderIn'],
  },
  {
    // the hyperlink is found whatever else (here an unknown element) sits inside cNvPr
    title: 'hyperlink next to unknown children',
    create() {
      return new CNvPrXform();
    },
    xml: `<xdr:cNvPr id="1" name="Picture 1"><a:hlinkClick xmlns:r="${R_NS}" r:id="rId1"/><a:hlinkHover/>${CREATION_ID}</xdr:cNvPr>`,
    parsedModel: { hyperlinks: { rId: 'rId1' } },
    tests: ['parse'],
  },
];

describe('CNvPrXform', () => {
  testXformHelper(expectations);
});
