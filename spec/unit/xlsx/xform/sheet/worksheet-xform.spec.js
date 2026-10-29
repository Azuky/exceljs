const fs = require('fs');
const { PassThrough } = require('readable-stream');

const testXformHelper = require('../test-xform-helper');

const Enums = verquire('doc/enums');
const RelType = verquire('xlsx/rel-type');
const parseSax = verquire('utils/parse-sax');
const { knownBug } = require('../../../../utils/known-bug');
const XmlStream = verquire('utils/xml-stream');
const WorksheetXform = verquire('xlsx/xform/sheet/worksheet-xform');

const SharedStringsXform = verquire('xlsx/xform/strings/shared-strings-xform');
const StylesXform = verquire('xlsx/xform/style/styles-xform');

const fakeStyles = {
  addStyleModel(style, cellType) {
    if (cellType === Enums.ValueType.Date) {
      return 1;
    }
    if (style && style.font) {
      return 2;
    }
    return 0;
  },
  getStyleModel(id) {
    switch (id) {
      case 1:
        return { numFmt: 'mm-dd-yy' };
      case 2:
        return {
          font: {
            underline: true,
            size: 11,
            color: { theme: 10 },
            name: 'Calibri',
            family: 2,
            scheme: 'minor',
          },
        };
      default:
        return null;
    }
  },
};

const fakeHyperlinkMap = {
  B6: 'https://www.npmjs.com/package/exceljs',
};

const fakeSharedStrings = {
  getString(index) {
    return ['Text', 'I am Text', 'Number', 'Integer', 'Date', 'Formula', 'Hyperlink', 'exceljs'][index];
  },
};

function fixDate(model) {
  model.rows[3].cells[1].value = new Date(model.rows[3].cells[1].value);
  return model;
}

// drop undefined members so models compare equal
function cloneForCompare(model) {
  return JSON.parse(JSON.stringify(model));
}

const expectations = [
  {
    title: 'Sheet 1',
    create: () => new WorksheetXform(),
    initialModel: fixDate(require('./data/sheet.1.0.json')),
    preparedModel: fixDate(require('./data/sheet.1.1.json')),
    xml: fs.readFileSync(`${__dirname}/data/sheet.1.2.xml`).toString(),
    parsedModel: require('./data/sheet.1.3.json'),
    reconciledModel: fixDate(require('./data/sheet.1.4.json')),
    tests: ['prepare', 'render', 'parse'],
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
    title: 'Sheet 1 - reconcile hyperlinks and styles',
    create: () => new WorksheetXform(),
    // the parsed sheet plus the part relationships that the reader attaches before reconciling
    parsedModel: {
      ...require('./data/sheet.1.3.json'),
      relationships: [
        {
          Id: 'rId1',
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
          Target: 'https://www.npmjs.com/package/exceljs',
          TargetMode: 'External',
        },
      ],
    },
    reconciledModel: fixDate(require('./data/sheet.1.4.json')),
    tests: ['reconcile'],
    options: {
      sharedStrings: fakeSharedStrings,
      hyperlinks: [],
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 2 - Data Validations',
    create: () => new WorksheetXform(),
    initialModel: require('./data/sheet.2.0.json'),
    preparedModel: require('./data/sheet.2.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/sheet.2.2.xml`).toString(),
    tests: ['prepare', 'render'],
    options: {
      styles: new StylesXform(true),
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 3 - Empty Sheet',
    create: () => new WorksheetXform(),
    preparedModel: require('./data/sheet.3.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/sheet.3.2.xml`).toString(),
    tests: ['render'],
    options: {
      styles: new StylesXform(true),
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
    },
  },
  {
    title: 'Sheet 5 - Shared Formulas',
    create: () => new WorksheetXform(),
    initialModel: require('./data/sheet.5.0.json'),
    preparedModel: require('./data/sheet.5.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/sheet.5.2.xml`).toString(),
    parsedModel: require('./data/sheet.5.3.json'),
    reconciledModel: require('./data/sheet.5.4.json'),
    tests: ['prepare-render', 'parse', 'reconcile'],
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
    title: 'Sheet 6 - AutoFilter',
    create: () => new WorksheetXform(),
    preparedModel: require('./data/sheet.6.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/sheet.6.2.xml`).toString(),
    parsedModel: require('./data/sheet.6.3.json'),
    tests: ['render', 'parse'],
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
    title: 'Sheet 7 - Row Breaks',
    create: () => new WorksheetXform(),
    initialModel: require('./data/sheet.7.0.json'),
    preparedModel: require('./data/sheet.7.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/sheet.7.2.xml`).toString(),
    tests: ['prepare', 'render'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
];

describe('WorksheetXform', () => {
  testXformHelper(expectations);

  it('hyperlinks must be after dataValidations', () => {
    const xform = new WorksheetXform();
    const model = require('./data/sheet.4.0.json');
    const xmlStream = new XmlStream();
    const options = {
      styles: new StylesXform(true),
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
    };
    xform.prepare(model, options);
    xform.render(xmlStream, model);

    const { xml } = xmlStream;
    const iHyperlinks = xml.indexOf('hyperlinks');
    const iDataValidations = xml.indexOf('dataValidations');
    expect(iHyperlinks).not.to.equal(-1);
    expect(iDataValidations).not.to.equal(-1);
    expect(iHyperlinks).to.be.greaterThan(iDataValidations);
  });

  it('conditionalFormattings must be before dataValidations', () => {
    const xform = new WorksheetXform();
    const model = require('./data/sheet.4.0.json');
    const xmlStream = new XmlStream();
    const options = {
      styles: new StylesXform(true),
      hyperlinks: [],
    };
    xform.prepare(model, options);
    xform.render(xmlStream, model);

    const { xml } = xmlStream;
    const iConditionalFormatting = xml.indexOf('conditionalFormatting');
    const iDataValidations = xml.indexOf('dataValidations');
    expect(iConditionalFormatting).not.to.equal(-1);
    expect(iDataValidations).not.to.equal(-1);
    expect(iConditionalFormatting).to.be.lessThan(iDataValidations);
  });

  describe('prepare relationships', () => {
    const emptySheet = () => ({
      cols: [],
      rows: [],
      media: [],
      tables: [],
      conditionalFormattings: [],
    });
    const prepareOptions = (extra) => ({
      styles: fakeStyles,
      sharedStrings: new SharedStringsXform(),
      ...extra,
    });

    it('adds an external hyperlink relationship for an image with a hyperlink', () => {
      const model = {
        ...emptySheet(),
        media: [
          {
            type: 'image',
            imageId: 0,
            range: { tl: { col: 0, row: 0 }, ext: { width: 100, height: 100 } },
            hyperlinks: { hyperlink: 'https://example.com/', tooltip: 'Example' },
          },
        ],
      };
      const options = prepareOptions({
        media: [{ name: 'image1', extension: 'png' }],
        drawings: [],
        drawingsCount: 0,
      });
      new WorksheetXform().prepare(model, options);

      // sheet -> drawing part
      expect(model.rels).to.deep.equal([
        {
          Id: 'rId1',
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          Target: '../drawings/drawing1.xml',
        },
      ]);
      expect(model.drawing.rId).to.equal('rId1');
      // drawing -> image and drawing -> external hyperlink
      expect(model.drawing.rels).to.deep.equal([
        { Id: 'rId1', Type: RelType.Image, Target: '../media/image1.png' },
        { Id: 'rId2', Type: RelType.Hyperlink, Target: 'https://example.com/', TargetMode: 'External' },
      ]);
      expect(model.drawing.anchors[0].picture).to.deep.equal({
        rId: 'rId1',
        hyperlinks: { tooltip: 'Example', rId: 'rId2' },
      });
    });

    it('adds a drawing relationship without hyperlink rel for a plain image', () => {
      const model = {
        ...emptySheet(),
        media: [{ type: 'image', imageId: 0, range: { tl: { col: 0, row: 0 }, ext: { width: 1, height: 1 } } }],
      };
      new WorksheetXform().prepare(
        model,
        prepareOptions({ media: [{ name: 'image1', extension: 'jpeg' }], drawings: [], drawingsCount: 0 }),
      );
      expect(model.drawing.rels).to.deep.equal([{ Id: 'rId1', Type: RelType.Image, Target: '../media/image1.jpeg' }]);
      expect(model.drawing.anchors[0].picture).to.deep.equal({ rId: 'rId1' });
    });

    it('relates tables and assigns a dxfId to styled table columns only', () => {
      const dxfStyles = [];
      const model = {
        ...emptySheet(),
        tables: [
          {
            target: 'table1.xml',
            columns: [{ name: 'Plain' }, { name: 'Bold', style: { font: { bold: true } } }],
          },
        ],
      };
      const options = prepareOptions({
        styles: {
          ...fakeStyles,
          addDxfStyle(style) {
            dxfStyles.push(style);
            return dxfStyles.length + 6;
          },
        },
      });
      new WorksheetXform().prepare(model, options);

      expect(model.rels).to.deep.equal([{ Id: 'rId1', Type: RelType.Table, Target: '../tables/table1.xml' }]);
      expect(model.tables[0].rId).to.equal('rId1');
      expect(model.tables[0].columns[0].dxfId).to.equal(undefined);
      expect(model.tables[0].columns[1].dxfId).to.equal(7);
      expect(dxfStyles).to.deep.equal([{ font: { bold: true } }]);
    });

    it('relates the pivot table part when the sheet has pivot tables', () => {
      const model = { ...emptySheet(), pivotTables: [{}] };
      new WorksheetXform().prepare(model, prepareOptions());
      expect(model.rels).to.deep.equal([
        { Id: 'rId1', Type: RelType.PivotTable, Target: '../pivotTables/pivotTable1.xml' },
      ]);
    });

    it('does not add a pivot relationship when there are no pivot tables', () => {
      const model = { ...emptySheet(), pivotTables: [] };
      new WorksheetXform().prepare(model, prepareOptions());
      expect(model.rels).to.deep.equal([]);
    });
  });

  describe('conditional formatting extensions', () => {
    const X14_URI = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
    const SHEET_OPEN =
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    const parseSheet = (xml) =>
      new Promise((resolve, reject) => {
        const stream = new PassThrough();
        stream.write(xml);
        stream.end();
        new WorksheetXform().parse(parseSax(stream)).then(resolve).catch(reject);
      });

    const extList = (cfs) =>
      `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14_URI}">` +
      `<x14:conditionalFormattings>${cfs}</x14:conditionalFormattings></ext></extLst>`;
    const extCf = (sqref, rules) =>
      `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${rules}<xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`;

    const dataBarRule = (id) =>
      '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>' +
      `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14_URI}"><x14:id>${id}</x14:id></ext></extLst></cfRule>`;
    const dataBarExtRule = (id) =>
      `<x14:cfRule type="dataBar" id="${id}"><x14:dataBar minLength="0" maxLength="100" gradient="0" border="1">` +
      '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF638EC6"/></x14:dataBar></x14:cfRule>';
    const iconExtRule = (id) =>
      `<x14:cfRule type="iconSet" priority="2" id="${id}"><x14:iconSet iconSet="3Triangles">` +
      '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo>' +
      '<x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule>';
    const cellIsRule =
      '<cfRule type="cellIs" dxfId="0" priority="3" operator="greaterThan"><formula>5</formula></cfRule>';

    it('merges an x14 data bar into the primary rule with the same x14Id', async () => {
      const id = '{AAAAAAAA-0000-0000-0000-000000000001}';
      const model = await parseSheet(
        `${SHEET_OPEN}<sheetData/><conditionalFormatting sqref="A1:A5">${dataBarRule(id)}</conditionalFormatting>` +
          `${extList(extCf('A1:A5', dataBarExtRule(id)))}</worksheet>`,
      );
      expect(model.conditionalFormattings).to.have.length(1);
      expect(model.conditionalFormattings[0].ref).to.equal('A1:A5');
      expect(model.conditionalFormattings[0].rules).to.have.length(1);
      const [rule] = model.conditionalFormattings[0].rules;
      expect(rule.type).to.equal('dataBar');
      expect(rule.priority).to.equal(1);
      expect(rule.x14Id).to.equal(id);
      // from the primary rule
      expect(rule.color).to.deep.equal({ argb: 'FF638EC6' });
      // from the extension: filled in because the primary rule does not define them
      expect(rule.gradient).to.equal(false);
      expect(rule.border).to.equal(true);
      expect(rule.borderColor).to.deep.equal({ argb: 'FF638EC6' });
      expect(rule.minLength).to.equal(0);
      expect(rule.maxLength).to.equal(100);
      // the primary rule keeps its own cfvo
      expect(cloneForCompare(rule.cfvo)).to.deep.equal([{ type: 'min' }, { type: 'max' }]);
    });

    it('adds an unmatched extension rule to the existing conditional formatting of the same ref', async () => {
      const model = await parseSheet(
        `${SHEET_OPEN}<sheetData/><conditionalFormatting sqref="A1:A5">${cellIsRule}</conditionalFormatting>` +
          `${extList(extCf('A1:A5', iconExtRule('{AAAAAAAA-0000-0000-0000-000000000002}')))}</worksheet>`,
      );
      expect(model.conditionalFormattings).to.have.length(1);
      const { rules } = model.conditionalFormattings[0];
      expect(rules.map((rule) => rule.type)).to.deep.equal(['cellIs', 'iconSet']);
      expect(rules[1].iconSet).to.equal('3Triangles');
      expect(rules[1].cfvo).to.deep.equal([
        { type: 'percent', value: 0 },
        { type: 'percent', value: 33 },
        { type: 'percent', value: 67 },
      ]);
    });

    it('appends an extension-only conditional formatting for a ref that has no primary rules', async () => {
      const model = await parseSheet(
        `${SHEET_OPEN}<sheetData/><conditionalFormatting sqref="A1:A5">${cellIsRule}</conditionalFormatting>` +
          `${extList(extCf('C1:C5', iconExtRule('{AAAAAAAA-0000-0000-0000-000000000003}')))}</worksheet>`,
      );
      expect(model.conditionalFormattings.map((cf) => cf.ref)).to.deep.equal(['A1:A5', 'C1:C5']);
      expect(model.conditionalFormattings[0].rules.map((rule) => rule.type)).to.deep.equal(['cellIs']);
      expect(model.conditionalFormattings[1].rules.map((rule) => rule.type)).to.deep.equal(['iconSet']);
    });

    it('returns the extension model when the sheet has no primary conditional formatting', async () => {
      const model = await parseSheet(
        `${SHEET_OPEN}<sheetData/>${extList(extCf('B2:B9', iconExtRule('{AAAAAAAA-0000-0000-0000-000000000004}')))}</worksheet>`,
      );
      expect(model.conditionalFormattings).to.have.length(1);
      expect(model.conditionalFormattings[0].ref).to.equal('B2:B9');
      expect(model.conditionalFormattings[0].rules[0].type).to.equal('iconSet');
    });

    it('keeps the primary model when the extension list is empty', async () => {
      const model = await parseSheet(
        `${SHEET_OPEN}<sheetData/><conditionalFormatting sqref="A1:A5">${cellIsRule}</conditionalFormatting></worksheet>`,
      );
      expect(model.conditionalFormattings.map((cf) => cf.ref)).to.deep.equal(['A1:A5']);
    });

    describe('write then read', () => {
      const writeThenRead = async (conditionalFormattings) => {
        const model = {
          ...require('./data/sheet.3.1.json'),
          conditionalFormattings,
        };
        const xform = new WorksheetXform();
        xform.prepare(model, {
          styles: new StylesXform(true),
          sharedStrings: new SharedStringsXform(),
          hyperlinks: [],
        });
        const xmlStream = new XmlStream();
        xform.render(xmlStream, model);
        return parseSheet(xmlStream.xml);
      };

      it('keeps a flat data bar with border and separate negative colors', async () => {
        const read = await writeThenRead([
          {
            ref: 'A1:A5',
            rules: [
              {
                type: 'dataBar',
                cfvo: [
                  { type: 'num', value: 5 },
                  { type: 'num', value: 20 },
                ],
                color: { argb: 'FF638EC6' },
                gradient: false,
                border: true,
                negativeBarColorSameAsPositive: false,
              },
            ],
          },
        ]);
        expect(read.conditionalFormattings).to.have.length(1);
        const [cf] = read.conditionalFormattings;
        expect(cf.ref).to.equal('A1:A5');
        expect(cf.rules).to.have.length(1);
        const rule = cf.rules[0];
        expect(rule.x14Id).to.match(/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/);
        delete rule.x14Id;
        expect(cloneForCompare(rule)).to.deep.equal({
          type: 'dataBar',
          priority: 1,
          cfvo: [
            { type: 'num', value: 5 },
            { type: 'num', value: 20 },
          ],
          color: { argb: 'FF638EC6' },
          minLength: 0,
          maxLength: 100,
          gradient: false,
          border: true,
          negativeBarColorSameAsPositive: false,
          negativeBarBorderColorSameAsPositive: true,
          axisPosition: 'auto',
          direction: 'leftToRight',
        });
      });

      it('keeps a 3Triangles icon set with custom icons', async () => {
        const read = await writeThenRead([
          {
            ref: 'B1:B9',
            rules: [
              {
                type: 'iconSet',
                iconSet: '3Triangles',
                priority: 4,
                cfvo: [
                  { type: 'percent', value: 0 },
                  { type: 'percent', value: 33 },
                  { type: 'percent', value: 67 },
                ],
                icons: [
                  { iconSet: '3Arrows', iconId: 0 },
                  { iconSet: '3Flags', iconId: 1 },
                  { iconSet: '3Symbols', iconId: 2 },
                ],
              },
            ],
          },
        ]);
        expect(read.conditionalFormattings).to.have.length(1);
        expect(read.conditionalFormattings[0].ref).to.equal('B1:B9');
        const rule = read.conditionalFormattings[0].rules[0];
        delete rule.x14Id;
        expect(cloneForCompare(rule)).to.deep.equal({
          type: 'iconSet',
          priority: 4,
          iconSet: '3Triangles',
          reverse: false,
          showValue: true,
          cfvo: [
            { type: 'percent', value: 0 },
            { type: 'percent', value: 33 },
            { type: 'percent', value: 67 },
          ],
          icons: [
            { iconSet: '3Arrows', iconId: 0 },
            { iconSet: '3Flags', iconId: 1 },
            { iconSet: '3Symbols', iconId: 2 },
          ],
        });
      });
    });
  });

  describe('element order', () => {
    // CT_Worksheet (ECMA-376 18.3.1.99): ... drawing, legacyDrawing, ..., picture, ..., tableParts, extLst
    knownBug(
      'legacyDrawing must come before tableParts',
      () => {
        const xform = new WorksheetXform();
        const xmlStream = new XmlStream();
        xform.render(xmlStream, {
          ...require('./data/sheet.3.1.json'),
          rels: [{ Id: 'rId1', Type: RelType.VmlDrawing, Target: '../drawings/vmlDrawing1.vml' }],
          tables: [{ rId: 'rId2' }],
        });
        const { xml } = xmlStream;
        const iLegacy = xml.indexOf('<legacyDrawing');
        const iTableParts = xml.indexOf('<tableParts');
        expect(iLegacy).not.to.equal(-1);
        expect(iTableParts).not.to.equal(-1);
        expect(iLegacy).to.be.lessThan(iTableParts);
      },
      /expected \d+ to be below \d+/,
    );
  });
});
