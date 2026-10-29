const testXformHelper = require('../test-xform-helper');

const TwoCellAnchorXform = verquire('xlsx/xform/drawing/two-cell-anchor-xform');

const picture = {
  index: 1,
  rId: 'rId1',
  hyperlinks: { tooltip: 'exceljs', rId: 'rId3' },
};
const range = {
  tl: { nativeRow: 0, nativeRowOff: 0, nativeCol: 0, nativeColOff: 0 },
  br: { nativeRow: 7, nativeRowOff: 0, nativeCol: 3, nativeColOff: 0 },
  editAs: 'oneCell',
};
const medium = { type: 'image', name: 'logo', extension: 'jpeg' };

const expectations = [
  {
    title: 'Picture anchored between two cells',
    create() {
      return new TwoCellAnchorXform();
    },
    initialModel: { range, picture: { rId: 'rId1', hyperlinks: picture.hyperlinks } },
    preparedModel: { range, picture },
    xml: `<xdr:twoCellAnchor editAs="oneCell">
      <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
      <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
      <xdr:pic>
        <xdr:nvPicPr>
          <xdr:cNvPr id="1" name="Picture 1">
            <a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId3" tooltip="exceljs"/>
            <a:extLst>
              <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{00000000-0008-0000-0000-000002000000}"/>
              </a:ext>
            </a:extLst>
          </xdr:cNvPr>
          <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
        </xdr:nvPicPr>
        <xdr:blipFill>
          <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print"/>
          <a:stretch><a:fillRect/></a:stretch>
        </xdr:blipFill>
        <xdr:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </xdr:spPr>
      </xdr:pic>
      <xdr:clientData/>
    </xdr:twoCellAnchor>`,
    parsedModel: { range, picture: { rId: 'rId1', hyperlinks: picture.hyperlinks } },
    reconciledModel: {
      range,
      picture: { rId: 'rId1', hyperlinks: picture.hyperlinks },
      medium,
    },
    options: {
      index: 0,
      rels: { rId1: { Target: '../media/image1.jpg' } },
      mediaIndex: { 'image1.jpg': 0 },
      media: [medium],
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
  },
];

describe('TwoCellAnchorXform', () => {
  testXformHelper(expectations);

  describe('reconcile', () => {
    // an anchor without a resolvable picture has no medium; the position data stays untouched
    it('leaves a null picture alone and resolves no medium', () => {
      const twoCell = new TwoCellAnchorXform();
      const model = { picture: null };
      twoCell.reconcile(model, {});
      expect(model.picture).to.equal(null);
      expect(model.medium).to.equal(undefined);
    });
    it('keeps br untouched when tl is missing', () => {
      const twoCell = new TwoCellAnchorXform();
      const model = { br: { col: 1, row: 1 } };
      twoCell.reconcile(model, {});
      expect(model.br).to.deep.equal({ col: 1, row: 1 });
      expect(model.tl).to.equal(undefined);
      expect(model.medium).to.equal(undefined);
    });
    it('keeps tl untouched when br is missing', () => {
      const twoCell = new TwoCellAnchorXform();
      const model = { tl: { col: 1, row: 1 } };
      twoCell.reconcile(model, {});
      expect(model.tl).to.deep.equal({ col: 1, row: 1 });
      expect(model.br).to.equal(undefined);
      expect(model.medium).to.equal(undefined);
    });
  });
});
