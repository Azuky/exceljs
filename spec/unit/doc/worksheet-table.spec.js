const { knownBug } = require('../../utils/known-bug');

const colCache = verquire('utils/col-cache');
const Excel = verquire('exceljs');
const Table = verquire('doc/table');

const spliceArray = (a, index, count, ...rest) => {
  const clone = [...a];
  clone.splice(index, count, ...rest);
  return clone;
};

const values = [
  ['Date', 'Id', 'Word'],
  [new Date('2019-08-01'), 1, 'Bird'],
  [new Date('2019-08-02'), 2, 'is'],
  [new Date('2019-08-03'), 3, 'the'],
  [new Date('2019-08-04'), 4, 'Word'],
  ['Totals', { formula: 'SUBTOTAL(104,TestTable[Id])', result: 4 }, null],
];

function addTable(ref, ws) {
  return ws.addTable({
    name: 'TestTable',
    ref,
    headerRow: true,
    totalsRow: true,
    style: {
      theme: 'TableStyleDark3',
      showRowStripes: true,
    },
    columns: [
      { name: 'Date', totalsRowLabel: 'Totals', filterButton: true },
      {
        name: 'Id',
        totalsRowFunction: 'max',
        filterButton: true,
        totalsRowResult: 4,
      },
      {
        name: 'Word',
        filterButton: false,
        style: { font: { bold: true, name: 'Comic Sans MS' } },
      },
    ],
    rows: [
      [new Date('2019-08-01'), 1, 'Bird'],
      [new Date('2019-08-02'), 2, 'is'],
      [new Date('2019-08-03'), 3, 'the'],
      [new Date('2019-08-04'), 4, 'Word'],
    ],
  });
}

function checkTable(ref, ws, testValues) {
  const a = colCache.decodeAddress(ref);

  for (let i = -1; i <= testValues.length + 1; i++) {
    const vRow = testValues[i];
    const nRow = i + a.row;
    const row = nRow >= 1 && ws.getRow(nRow);
    if (!row) continue;
    for (let j = -1; j <= testValues[0].length + 1; j++) {
      const value = (vRow && vRow[j]) || null;
      const nCol = j + a.col;
      if (nCol < 1) continue;
      const cellValue = row.getCell(nCol).value;

      if (value instanceof Date) {
        expect(cellValue).to.equalDate(value);
      } else if (value === null) {
        expect(cellValue).to.be.null;
      } else if (typeof value === 'object') {
        expect(cellValue).to.deep.equal(value);
      } else {
        expect(cellValue).to.equal(value);
      }
    }
  }
}

describe('Worksheet', () => {
  describe('Table', () => {
    it('creates a table', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      addTable('A1', ws);

      checkTable('A1', ws, values);
    });

    it('removes header', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.headerRow = false;
      table.commit();

      const newValues = spliceArray(values, 0, 1);
      checkTable('A1', ws, newValues);
    });

    it('removes totals', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.totalsRow = false;
      table.commit();

      const newValues = spliceArray(values, 5, 1);
      checkTable('A1', ws, newValues);
    });

    it('moves the table', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.ref = 'C2';
      table.commit();

      checkTable('C2', ws, values);
    });

    it('removes a row', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.removeRows(1);
      table.commit();

      const newValues = spliceArray(values, 2, 1);
      checkTable('A1', ws, newValues);
    });

    it('adds a row', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.addRow([new Date('2019-08-05'), 5, 'Bird']);
      table.commit();

      const newValues = spliceArray(values, 5, 0, [new Date('2019-08-05'), 5, 'Bird']);
      checkTable('A1', ws, newValues);
    });

    it('removes a column', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.removeColumns(1);
      table.commit();

      const newValues = values.map((rVals) => spliceArray(rVals, 1, 1));
      checkTable('A1', ws, newValues);
    });

    it('adds a column', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      table.addColumn(
        {
          name: 'Letter',
          totalsRowFunction: 'custom',
          totalsRowFormula: 'ROW()',
          totalsRowResult: 6,
          filterButton: true,
        },
        ['a', 'b', 'c', 'd'],
        2,
      );
      table.commit();

      const colValues = ['Letter', 'a', 'b', 'c', 'd', { formula: 'ROW()', result: 6 }];
      const newValues = values.map((rVals, i) => spliceArray(rVals, 2, 0, colValues[i]));
      checkTable('A1', ws, newValues);
    });

    it('renames a column', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);

      const column = table.getColumn(1);
      column.name = 'Code';
      table.commit();

      const newValues = [...values];
      newValues.splice(0, 1, ['Date', 'Code', 'Word']);
      newValues.splice(5, 1, ['Totals', { formula: 'SUBTOTAL(104,TestTable[Code])', result: 4 }, null]);

      checkTable('A1', ws, newValues);
    });
    describe('displayName', () => {
      it('defaults to the table name', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        const table = addTable('A1', ws);
        expect(table.displayName).to.equal('TestTable');
      });

      // known bug: the setter is declared as `displayNamename` and the getter reads `displyName`, so assigning displayName has no effect — see stage-2 notes
      knownBug(
        'stores a new display name',
        () => {
          const ws = new Excel.Workbook().addWorksheet('blort');
          const table = addTable('A1', ws);
          table.displayName = 'Renamed';
          table.commit();
          expect(table.displayName).to.equal('Renamed');
        },
        /only a getter/,
      );
    });

    describe('registry', () => {
      function twoTables(ws) {
        const t1 = ws.addTable({ name: 't1', ref: 'A1', columns: [{ name: 'a' }], rows: [[1]] });
        const t2 = ws.addTable({ name: 't2', ref: 'D1', columns: [{ name: 'b' }], rows: [[2]] });
        return [t1, t2];
      }

      it('getTable and getTables return the added tables', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        const [t1, t2] = twoTables(ws);
        expect(ws.getTables()).to.have.length(2);
        expect(ws.getTable('t1')).to.equal(t1);
        expect(ws.getTable('t2')).to.equal(t2);
        expect(ws.getTable('missing')).to.equal(undefined);
      });

      it('removeTable drops the table by name', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        const [, t2] = twoTables(ws);
        ws.removeTable('t1');
        expect(ws.getTables()).to.deep.equal([t2]);
        expect(ws.getTable('t1')).to.equal(undefined);
      });

      it('removeTable ignores an unknown name', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        twoTables(ws);
        expect(() => ws.removeTable('missing')).not.to.throw();
        expect(ws.getTables()).to.have.length(2);
      });
    });

    describe('totals row function', () => {
      const expectations = {
        average: 'SUBTOTAL(101,Totals[n])',
        countNums: 'SUBTOTAL(102,Totals[n])',
        count: 'SUBTOTAL(103,Totals[n])',
        max: 'SUBTOTAL(104,Totals[n])',
        min: 'SUBTOTAL(105,Totals[n])',
        stdDev: 'SUBTOTAL(106,Totals[n])',
        var: 'SUBTOTAL(107,Totals[n])',
        sum: 'SUBTOTAL(109,Totals[n])',
      };

      function totalsTable(ws, column) {
        return ws.addTable({
          name: 'Totals',
          ref: 'A1',
          totalsRow: true,
          columns: [{ name: 'label' }, { name: 'n', ...column }],
          rows: [
            ['x', 1],
            ['y', 2],
          ],
        });
      }

      Object.keys(expectations).forEach((totalsRowFunction) => {
        it(`writes the SUBTOTAL formula for ${totalsRowFunction}`, () => {
          const ws = new Excel.Workbook().addWorksheet('blort');
          totalsTable(ws, { totalsRowFunction, totalsRowResult: 3 });
          expect(ws.getCell('B4').value).to.deep.equal({ formula: expectations[totalsRowFunction], result: 3 });
        });
      });

      it('writes the given formula for a custom function', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        totalsTable(ws, { totalsRowFunction: 'custom', totalsRowFormula: 'ROW()', totalsRowResult: 4 });
        expect(ws.getCell('B4').value).to.deep.equal({ formula: 'ROW()', result: 4 });
      });

      it('writes an empty cell for the none function', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        totalsTable(ws, { totalsRowFunction: 'none' });
        expect(ws.getCell('B4').value).to.equal(null);
      });

      it('defaults to none when no function is given and labels the first column', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        totalsTable(ws, {});
        expect(ws.getCell('B4').value).to.equal(null);
        expect(ws.getCell('A4').value).to.equal('Total');
      });

      it('rejects an unknown function', () => {
        const ws = new Excel.Workbook().addWorksheet('blort');
        expect(() => totalsTable(ws, { totalsRowFunction: 'bogus' })).to.throw('Invalid Totals Row Function: bogus');
      });
    });

    describe('load', () => {
      it('writes headers, rows and totals of a table model into a sheet', () => {
        const wb = new Excel.Workbook();
        const source = addTable('B2', wb.addWorksheet('source'));
        const target = wb.addWorksheet('target');

        const table = new Table(target);
        table.model = source.model;
        table.load(target);

        checkTable('B2', target, values);
        expect(target.getCell('B2').value).to.equal('Date');
        expect(target.getCell('D6').value).to.equal('Word');
        expect(target.getCell('C7').value).to.deep.equal({ formula: 'SUBTOTAL(104,TestTable[Id])', result: 4 });
      });

      it('omits the header and totals rows when the model has none', () => {
        const wb = new Excel.Workbook();
        const source = addTable('A1', wb.addWorksheet('source'));
        source.model.headerRow = false;
        source.model.totalsRow = false;
        const target = wb.addWorksheet('target');

        const table = new Table(target);
        table.model = source.model;
        table.load(target);

        expect(target.rowCount).to.equal(4);
        expect(target.getCell('A1').value).to.equalDate(new Date('2019-08-01'));
      });
    });
  });
});
