const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const SharedStringsXform = verquire('xlsx/xform/strings/shared-strings-xform');

const expectations = [
  {
    title: 'Shared Strings',
    create() {
      return new SharedStringsXform();
    },
    preparedModel: require('./data/sharedStrings.json'),
    xml: fs.readFileSync(`${__dirname}/data/sharedStrings.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('SharedStringsXform', () => {
  testXformHelper(expectations);

  describe('unexpected nodes', () => {
    it('rejects an unknown child in parseOpen', async () => {
      let error;
      try {
        await testXformHelper.parseXml(new SharedStringsXform(), '<sst><bogus/></sst>');
      } catch (e) {
        error = e;
      }
      expect(error).to.be.an('error');
      expect(error.message).to.match(/^Unexpected xml node in parseOpen: \{"name":"bogus","attributes":\{\}/);
    });

    it('throws on an unexpected closing tag', () => {
      expect(() => new SharedStringsXform().parseClose('bogus')).to.throw('Unexpected xml node in parseClose: bogus');
    });
  });
});
