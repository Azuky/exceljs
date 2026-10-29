const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const RelationshipsXform = verquire('xlsx/xform/core/relationships-xform');

const expectations = [
  {
    title: 'worksheet.rels',
    create() {
      return new RelationshipsXform();
    },
    preparedModel: require('./data/worksheet.rels.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/worksheet.rels.xml`).toString().replace(/\r\n/g, '\n'),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('RelationshipsXform', () => {
  testXformHelper(expectations);

  describe('unexpected nodes', () => {
    it('rejects an unknown child in parseOpen', async () => {
      let error;
      try {
        await testXformHelper.parseXml(new RelationshipsXform(), '<Relationships><bogus/></Relationships>');
      } catch (e) {
        error = e;
      }
      expect(error).to.be.an('error');
      expect(error.message).to.match(/^Unexpected xml node in parseOpen: \{"name":"bogus","attributes":\{\}/);
    });

    it('throws on an unexpected closing tag', () => {
      expect(() => new RelationshipsXform().parseClose('bogus')).to.throw('Unexpected xml node in parseClose: bogus');
    });
  });
});
