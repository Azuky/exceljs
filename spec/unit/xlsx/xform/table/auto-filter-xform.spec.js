const testXformHelper = require('../test-xform-helper');

const AutoFilterXform = verquire('xlsx/xform/table/auto-filter-xform');

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new AutoFilterXform();
    },
    initialModel: {
      autoFilterRef: 'A1:B10',
      columns: [{ filterButton: false }, { filterButton: true }, { filterButton: true }],
    },
    preparedModel: {
      autoFilterRef: 'A1:B10',
      columns: [
        { colId: '0', filterButton: false },
        { colId: '1', filterButton: true },
        { colId: '2', filterButton: true },
      ],
    },
    xml:
      '<autoFilter ref="A1:B10">' +
      '<filterColumn colId="0" hiddenButton="1" />' +
      '<filterColumn colId="1" hiddenButton="0" />' +
      '<filterColumn colId="2" hiddenButton="0" />' +
      '</autoFilter>',
    get parsedModel() {
      return this.initialModel;
    },
    tests: ['prepare', 'render', 'renderIn', 'parse'],
  },
];

describe('AutoFilterXform', () => {
  testXformHelper(expectations);

  describe('unexpected nodes', () => {
    it('rejects an unknown child in parseOpen', async () => {
      let error;
      try {
        await testXformHelper.parseXml(new AutoFilterXform(), '<autoFilter ref="A1:B2"><bogus/></autoFilter>');
      } catch (e) {
        error = e;
      }
      expect(error).to.be.an('error');
      expect(error.message).to.match(/^Unexpected xml node in parseOpen: \{"name":"bogus","attributes":\{\}/);
    });

    it('throws on an unexpected closing tag', () => {
      expect(() => new AutoFilterXform().parseClose('bogus')).to.throw('Unexpected xml node in parseClose: bogus');
    });
  });
});
