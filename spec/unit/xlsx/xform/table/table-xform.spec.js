const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const TableXform = verquire('xlsx/xform/table/table-xform');
const TableColumnXform = verquire('xlsx/xform/table/table-column-xform');
const { knownBug } = require('../../../../utils/known-bug');

const preparedTable = require('./data/table.1.1');

// ids are assigned by prepare: strip them from a copy of the prepared model
const initialTable = JSON.parse(JSON.stringify(preparedTable));
initialTable.columns.forEach((column) => {
  delete column.id;
  delete column.colId;
});

const dxfStyle = { font: { bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } };

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new TableXform();
    },
    initialModel: initialTable,
    options: {},
    preparedModel: preparedTable,
    xml: fs.readFileSync(`${__dirname}/data/table.1.2.xml`).toString(),
    parsedModel: require('./data/table.1.3'),
    tests: ['prepare', 'render', 'renderIn', 'parse'],
  },
  {
    title: 'column with dxf style',
    create() {
      return new TableXform();
    },
    preparedModel: {
      id: 1,
      name: 'Styled',
      displayName: 'Styled',
      tableRef: 'A1:B3',
      autoFilterRef: 'A1:B3',
      totalsRow: false,
      headerRow: true,
      columns: [
        { id: 1, colId: '0', filterButton: true, name: 'Plain' },
        { id: 2, colId: '1', filterButton: true, name: 'Bold', dxfId: 0 },
      ],
      style: {
        theme: 'TableStyleMedium2',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    },
    xml:
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" ' +
      'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" ' +
      'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" ' +
      'id="1" name="Styled" displayName="Styled" ref="A1:B3" totalsRowShown="1" headerRowCount="1">' +
      '<autoFilter ref="A1:B3"><filterColumn colId="0" hiddenButton="0"/><filterColumn colId="1" hiddenButton="0"/></autoFilter>' +
      '<tableColumns count="2"><tableColumn id="1" name="Plain"/><tableColumn id="2" name="Bold" dataDxfId="0"/></tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>' +
      '</table>',
    // sax delivers attribute values as strings
    parsedModel: {
      name: 'Styled',
      displayName: 'Styled',
      tableRef: 'A1:B3',
      autoFilterRef: 'A1:B3',
      totalsRow: false,
      headerRow: true,
      columns: [
        { filterButton: true, name: 'Plain' },
        { filterButton: true, name: 'Bold', dxfId: '0' },
      ],
      style: {
        theme: 'TableStyleMedium2',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    },
    // reconcile resolves dxfId against the styles part and leaves plain columns alone
    reconciledModel: {
      name: 'Styled',
      displayName: 'Styled',
      tableRef: 'A1:B3',
      autoFilterRef: 'A1:B3',
      totalsRow: false,
      headerRow: true,
      columns: [
        { filterButton: true, name: 'Plain' },
        { filterButton: true, name: 'Bold', dxfId: '0', style: dxfStyle },
      ],
      style: {
        theme: 'TableStyleMedium2',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    },
    options: {
      styles: {
        getDxfStyle(id) {
          return id === '0' ? dxfStyle : undefined;
        },
      },
    },
    // known bug: lib writes/reads `dxfId`; CT_TableColumn only has dataDxfId (headerRowDxfId, totalsRowDxfId)
    knownFailures: ['render'],
    knownFailureMessage: /tableColumn\[1\]: attributes \[dxfId,id,name\] !== \[dataDxfId,id,name\]/,
    tests: ['render', 'reconcile'],
  },
];

describe('TableXform', () => {
  testXformHelper(expectations);

  // known bug: the column's differential format is read from `dxfId`, the attribute Excel writes is `dataDxfId`
  knownBug(
    'parses the dataDxfId of a table column',
    async () => {
      const model = await testXformHelper.parseXml(
        new TableColumnXform(),
        '<tableColumn id="2" name="Bold" dataDxfId="0"/>',
      );
      expect(model.dxfId).to.equal('0');
    },
    /expected undefined to equal '0'/,
  );
});
