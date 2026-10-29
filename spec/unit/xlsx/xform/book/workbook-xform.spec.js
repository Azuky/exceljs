const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const WorkbookXform = verquire('xlsx/xform/book/workbook-xform');

const expectations = [
  {
    title: 'book.1',
    create() {
      return new WorkbookXform();
    },
    preparedModel: require('./data/book.1.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/book.1.2.xml`).toString().replace(/\r\n/g, '\n'),
    parsedModel: require('./data/book.1.3.json'),
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'book.2 - no properties',
    create() {
      return new WorkbookXform();
    },
    xml: fs.readFileSync(`${__dirname}/data/book.2.2.xml`).toString().replace(/\r\n/g, '\n'),
    parsedModel: require('./data/book.2.3.json'),
    tests: ['parse'],
  },
];

describe('WorkbookXform', () => {
  testXformHelper(expectations);

  describe('reconcile', () => {
    const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const makeModel = (overrides) => {
      const worksheetHash = {
        'xl/worksheets/sheet1.xml': { id: 0 },
        'xl/worksheets/sheet2.xml': { id: 0 },
      };
      return {
        sheets: [
          { name: 'First', id: 1, state: 'visible', rId: 'rId1' },
          { name: 'Second', id: 2, state: 'hidden', rId: 'rId2' },
        ],
        workbookRels: [
          { Id: 'rId1', Type: `${REL_BASE}/worksheet`, Target: 'worksheets/sheet1.xml' },
          { Id: 'rId2', Type: `${REL_BASE}/worksheet`, Target: 'worksheets/sheet2.xml' },
        ],
        worksheetHash,
        definedNames: [],
        media: [],
        ...overrides,
      };
    };
    const reconcile = (model) => {
      new WorkbookXform().reconcile(model);
      return model;
    };

    it('copies name, id and state of each sheet to its worksheet', () => {
      const model = reconcile(makeModel());
      expect(model.worksheetHash['xl/worksheets/sheet1.xml']).to.deep.equal({ id: 1, name: 'First', state: 'visible' });
      expect(model.worksheetHash['xl/worksheets/sheet2.xml']).to.deep.equal({ id: 2, name: 'Second', state: 'hidden' });
    });

    it('resolves targets with a leading /xl/ or leading spaces', () => {
      const model = reconcile(
        makeModel({
          workbookRels: [
            { Id: 'rId1', Type: `${REL_BASE}/worksheet`, Target: '/xl/worksheets/sheet1.xml' },
            { Id: 'rId2', Type: `${REL_BASE}/worksheet`, Target: '  worksheets/sheet2.xml' },
          ],
        }),
      );
      expect(model.worksheetHash['xl/worksheets/sheet1.xml'].name).to.equal('First');
      expect(model.worksheetHash['xl/worksheets/sheet2.xml'].name).to.equal('Second');
    });

    it('ignores chartsheets and sheets without a relationship', () => {
      const model = reconcile(
        makeModel({
          sheets: [
            { name: 'Chart', id: 1, rId: 'rId3' },
            { name: 'Orphan', id: 2, rId: 'rId9' },
            { name: 'Data', id: 3, state: 'visible', rId: 'rId2' },
          ],
          workbookRels: [
            { Id: 'rId2', Type: `${REL_BASE}/worksheet`, Target: 'worksheets/sheet2.xml' },
            { Id: 'rId3', Type: `${REL_BASE}/chartsheet`, Target: 'chartsheets/sheet1.xml' },
          ],
        }),
      );
      expect(model.worksheetHash['xl/worksheets/sheet2.xml']).to.deep.equal({ id: 3, name: 'Data', state: 'visible' });
      expect(model.worksheetHash['xl/worksheets/sheet1.xml']).to.deep.equal({ id: 0 });
    });

    it('creates the page setup and sets the print area from _xlnm.Print_Area', () => {
      const model = reconcile(
        makeModel({
          definedNames: [{ name: '_xlnm.Print_Area', localSheetId: 1, ranges: ['Second!$B$2:$D$9'] }],
        }),
      );
      expect(model.worksheetHash['xl/worksheets/sheet2.xml'].pageSetup).to.deep.equal({ printArea: 'B2:D9' });
      expect(model.worksheetHash['xl/worksheets/sheet1.xml'].pageSetup).to.equal(undefined);
      expect(model.definedNames).to.deep.equal([]);
    });

    it('joins several print areas of the same sheet with &&', () => {
      const model = reconcile(
        makeModel({
          definedNames: [
            { name: '_xlnm.Print_Area', localSheetId: 0, ranges: ['First!$A$1:$B$2'] },
            { name: '_xlnm.Print_Area', localSheetId: 0, ranges: ['First!$D$4:$E$5'] },
          ],
        }),
      );
      expect(model.worksheetHash['xl/worksheets/sheet1.xml'].pageSetup.printArea).to.equal('A1:B2&&D4:E5');
    });

    it('keeps an existing page setup when adding the print area', () => {
      const model = makeModel({
        definedNames: [{ name: '_xlnm.Print_Area', localSheetId: 0, ranges: ['First!$A$1:$B$2'] }],
      });
      model.worksheetHash['xl/worksheets/sheet1.xml'].pageSetup = { orientation: 'landscape' };
      reconcile(model);
      expect(model.worksheetHash['xl/worksheets/sheet1.xml'].pageSetup).to.deep.equal({
        orientation: 'landscape',
        printArea: 'A1:B2',
      });
    });

    it('splits _xlnm.Print_Titles into rows only', () => {
      const model = reconcile(
        makeModel({
          definedNames: [{ name: '_xlnm.Print_Titles', localSheetId: 0, ranges: ['First!$1:$2'] }],
        }),
      );
      expect(model.worksheetHash['xl/worksheets/sheet1.xml'].pageSetup).to.deep.equal({ printTitlesRow: '1:2' });
    });

    it('splits _xlnm.Print_Titles into columns only', () => {
      const model = reconcile(
        makeModel({
          definedNames: [{ name: '_xlnm.Print_Titles', localSheetId: 0, ranges: ['First!$A:$B'] }],
        }),
      );
      expect(model.worksheetHash['xl/worksheets/sheet1.xml'].pageSetup).to.deep.equal({ printTitlesColumn: 'A:B' });
    });

    it('splits _xlnm.Print_Titles into columns and rows', () => {
      const model = reconcile(
        makeModel({
          definedNames: [{ name: '_xlnm.Print_Titles', localSheetId: 1, ranges: ['Second!$A:$C', 'Second!$3:$4'] }],
        }),
      );
      expect(model.worksheetHash['xl/worksheets/sheet2.xml'].pageSetup).to.deep.equal({
        printTitlesColumn: 'A:C',
        printTitlesRow: '3:4',
      });
    });

    it('drops print names that point to a missing sheet and keeps other names', () => {
      const other = { name: 'MyRange', ranges: ['First!$A$1'] };
      const scoped = { name: 'Local', localSheetId: 1, ranges: ['Second!$C$3'] };
      const model = reconcile(
        makeModel({
          definedNames: [
            { name: '_xlnm.Print_Area', localSheetId: 7, ranges: ['Nowhere!$A$1:$B$2'] },
            { name: '_xlnm.Print_Titles', localSheetId: 7, ranges: ['Nowhere!$1:$2'] },
            other,
            scoped,
          ],
        }),
      );
      expect(model.definedNames).to.deep.equal([other, scoped]);
      expect(model.worksheetHash['xl/worksheets/sheet1.xml'].pageSetup).to.equal(undefined);
      expect(model.worksheetHash['xl/worksheets/sheet2.xml'].pageSetup).to.equal(undefined);
    });

    it('numbers the media for the sheets', () => {
      const model = reconcile(makeModel({ media: [{}, {}, {}] }));
      expect(model.media.map((media) => media.index)).to.deep.equal([0, 1, 2]);
    });
  });
});
