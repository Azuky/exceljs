const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const ContentTypesXform = verquire('xlsx/xform/core/content-types-xform');

const expectations = [
  {
    title: 'Three Sheets with shared strings',
    create() {
      return new ContentTypesXform();
    },
    preparedModel: {
      worksheets: [{ id: 1 }, { id: 2 }, { id: 3 }],
      media: [],
      drawings: [],
      sharedStrings: { count: 1 },
    },
    xml: fs.readFileSync(`${__dirname}/data/content-types.01.xml`).toString().replace(/\r\n/g, '\n'),
    tests: ['render'],
  },
  {
    title: 'Images with shared strings',
    create() {
      return new ContentTypesXform();
    },
    preparedModel: {
      worksheets: [{ id: 1 }, { id: 2 }],
      media: [
        { type: 'image', extension: 'png' },
        { type: 'image', extension: 'jpg' },
      ],
      drawings: [],
      sharedStrings: { count: 1 },
    },
    xml: fs.readFileSync(`${__dirname}/data/content-types.02.xml`).toString().replace(/\r\n/g, '\n'),
    tests: ['render'],
  },
  {
    title: 'Three Sheets without shared strings',
    create() {
      return new ContentTypesXform();
    },
    preparedModel: {
      worksheets: [{ id: 1 }, { id: 2 }, { id: 3 }],
      media: [],
      drawings: [],
    },
    xml: fs.readFileSync(`${__dirname}/data/content-types.03.xml`).toString().replace(/\r\n/g, '\n'),
    tests: ['render'],
  },
  {
    title: 'Images without shared strings',
    create() {
      return new ContentTypesXform();
    },
    preparedModel: {
      worksheets: [{ id: 1 }, { id: 2, useSharedStrings: false }],
      media: [
        { type: 'image', extension: 'png' },
        { type: 'image', extension: 'jpg' },
      ],
      drawings: [],
    },
    xml: fs.readFileSync(`${__dirname}/data/content-types.04.xml`).toString().replace(/\r\n/g, '\n'),
    tests: ['render'],
  },
  {
    // parts other than the workbook and worksheets get an Override entry each;
    // vml drawings (comments) additionally need a Default entry for the .vml extension
    title: 'Tables, drawings, comments and a pivot table',
    create() {
      return new ContentTypesXform();
    },
    preparedModel: {
      worksheets: [{ id: 1 }],
      media: [{ type: 'image', extension: 'png' }],
      tables: [{ target: 'table1.xml' }, { target: 'table2.xml' }],
      drawings: [{ name: 'drawing1' }, { name: 'drawing2' }],
      commentRefs: [
        { commentName: 'comments1', vmlDrawing: 'vmlDrawing1' },
        { commentName: 'comments2', vmlDrawing: 'vmlDrawing2' },
      ],
      pivotTables: [{}],
      sharedStrings: { count: 1 },
    },
    xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/xl/comments2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`,
    tests: ['render'],
  },
];

describe('ContentTypesXform', () => {
  testXformHelper(expectations);
});
