const testXformHelper = require('../test-xform-helper');

const OneCellAnchorXform = verquire('xlsx/xform/drawing/one-cell-anchor-xform');

const medium = { type: 'image', name: 'logo', extension: 'png' };
const range = {
  tl: { nativeRow: 2, nativeRowOff: 90000, nativeCol: 5, nativeColOff: 320000 },
  ext: { width: 100, height: 200 },
  editAs: 'oneCell',
};

const expectations = [
  {
    title: 'Picture anchored at one cell with a size',
    create() {
      return new OneCellAnchorXform();
    },
    initialModel: { range, picture: { rId: 'rId2' } },
    preparedModel: { range, picture: { index: 2, rId: 'rId2' } },
    xml: `<xdr:oneCellAnchor editAs="oneCell">
      <xdr:from><xdr:col>5</xdr:col><xdr:colOff>320000</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>90000</xdr:rowOff></xdr:from>
      <xdr:ext cx="952500" cy="1905000"/>
      <xdr:pic>
        <xdr:nvPicPr>
          <xdr:cNvPr id="2" name="Picture 2">
            <a:extLst>
              <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{00000000-0008-0000-0000-000002000000}"/>
              </a:ext>
            </a:extLst>
          </xdr:cNvPr>
          <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
        </xdr:nvPicPr>
        <xdr:blipFill>
          <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2" cstate="print"/>
          <a:stretch><a:fillRect/></a:stretch>
        </xdr:blipFill>
        <xdr:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </xdr:spPr>
      </xdr:pic>
      <xdr:clientData/>
    </xdr:oneCellAnchor>`,
    parsedModel: { range, picture: { rId: 'rId2' } },
    reconciledModel: { range, picture: { rId: 'rId2' }, medium },
    options: {
      index: 1,
      rels: { rId2: { Target: '../media/image2.png' } },
      mediaIndex: { 'image2.png': 0 },
      media: [medium],
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
  },
];

describe('OneCellAnchorXform', () => {
  testXformHelper(expectations);
});
