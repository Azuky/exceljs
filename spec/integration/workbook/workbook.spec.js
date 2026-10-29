const { knownBug } = require('../../utils/known-bug');

const { PassThrough, Readable } = require('stream');

const testUtils = require('../../utils/index');

const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';
const TEST_CSV_FILE_NAME = './spec/out/wb.test.csv';

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Serialise', () => {
    it('xlsx file', () => {
      const wb = testUtils.createTestBook(new ExcelJS.Workbook(), 'xlsx');

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          testUtils.checkTestBook(wb2, 'xlsx');
        });
    });
    describe('Xlsx Zip Compression', () => {
      it('xlsx file with best compression', () => {
        const wb = testUtils.createTestBook(new ExcelJS.Workbook(), 'xlsx');

        return wb.xlsx
          .writeFile(TEST_XLSX_FILE_NAME, {
            zip: {
              compression: 'DEFLATE',
              compressionOptions: {
                level: 9,
              },
            },
          })
          .then(() => {
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
          })
          .then((wb2) => {
            testUtils.checkTestBook(wb2, 'xlsx');
          });
      });

      it('xlsx file with default compression', () => {
        const wb = testUtils.createTestBook(new ExcelJS.Workbook(), 'xlsx');

        return wb.xlsx
          .writeFile(TEST_XLSX_FILE_NAME, {
            zip: {
              compression: 'DEFLATE',
            },
          })
          .then(() => {
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
          })
          .then((wb2) => {
            testUtils.checkTestBook(wb2, 'xlsx');
          });
      });

      it('xlsx file with fast compression', () => {
        const wb = testUtils.createTestBook(new ExcelJS.Workbook(), 'xlsx');

        return wb.xlsx
          .writeFile(TEST_XLSX_FILE_NAME, {
            zip: {
              compression: 'DEFLATE',
              compressionOptions: {
                level: 1,
              },
            },
          })
          .then(() => {
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
          })
          .then((wb2) => {
            testUtils.checkTestBook(wb2, 'xlsx');
          });
      });

      it('xlsx file with no compression', () => {
        const wb = testUtils.createTestBook(new ExcelJS.Workbook(), 'xlsx');

        return wb.xlsx
          .writeFile(TEST_XLSX_FILE_NAME, {
            zip: {
              compression: 'STORE',
            },
          })
          .then(() => {
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
          })
          .then((wb2) => {
            testUtils.checkTestBook(wb2, 'xlsx');
          });
      });
    });
    it('sheets with correct names', () => {
      const wb = new ExcelJS.Workbook();
      const ws1 = wb.addWorksheet('Hello, World!');
      expect(ws1.name).to.equal('Hello, World!');
      ws1.getCell('A1').value = 'Hello, World!';

      const ws2 = wb.addWorksheet();
      expect(ws2.name).to.match(/sheet\d+/);
      ws2.getCell('A1').value = ws2.name;

      wb.addWorksheet('This & That');

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          expect(wb2.getWorksheet('Hello, World!')).to.be.ok;
          expect(wb2.getWorksheet('This & That')).to.be.ok;
        });
    });

    it('creator, lastModifiedBy, etc', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'World!';
      wb.creator = 'Foo';
      wb.lastModifiedBy = 'Bar';
      wb.created = new Date(2016, 0, 1);
      wb.modified = new Date(2016, 4, 19);
      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          expect(wb2.creator).to.equal(wb.creator);
          expect(wb2.lastModifiedBy).to.equal(wb.lastModifiedBy);
          expect(wb2.created).to.equalDate(wb.created);
          expect(wb2.modified).to.equalDate(wb.modified);
        });
    });
    it('printTitlesRow', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('printHeader');

      ws.getCell('A1').value = 'This is a header row repeated on every printed page';
      ws.getCell('B2').value = 'This is a header row too';

      for (let i = 0; i < 100; i++) {
        ws.addRow(['not header row']);
      }

      ws.pageSetup.printTitlesRow = '1:2';

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('printHeader');
          expect(ws2.pageSetup.printTitlesRow).to.equal('1:2');
          expect(ws2.pageSetup.printTitlesColumn).to.be.undefined;
        });
    });
    it('printTitlesColumn', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('printColumn');

      ws.getCell('A1').value = 'This is a column repeated on every printed page';
      ws.getCell('A2').value = 'This is a column repeated on every printed page';
      ws.getCell('B1').value = 'This is a repeated column too';
      ws.getCell('B2').value = 'This is a repeated column too';

      ws.getCell('C1').value = 'This is a regular column';
      ws.getCell('C2').value = 'This is a regular column';
      ws.getCell('D1').value = 'This is a regular column';
      ws.getCell('D2').value = 'This is a regular column';

      ws.pageSetup.printTitlesRow = 'A:B';

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('printColumn');
          expect(ws2.pageSetup.printTitlesRow).to.be.undefined;
          expect(ws2.pageSetup.printTitlesColumn).to.equal('A:B');
        });
    });
    it('printTitlesRowAndColumn', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('printHeaderAndColumn');

      ws.getCell('A1').value = 'This is a column / row repeated on every printed page';
      ws.getCell('A2').value = 'This is a column / row repeated on every printed page';
      ws.getCell('B1').value = 'This is a repeated column / row too';
      ws.getCell('B2').value = 'This is a repeated column / row too';

      ws.getCell('C1').value = 'This is a regular column, repeated row';
      ws.getCell('C2').value = 'This is a regular column, repeated row';
      ws.getCell('D1').value = 'This is a regular column, repeated row';
      ws.getCell('D2').value = 'This is a regular column, repeated row';

      ws.getCell('A3').value = 'This is a repeated column';
      ws.getCell('B3').value = 'This is a repeated column';
      ws.getCell('C3').value = 'This is a regular column / row';
      ws.getCell('D3').value = 'This is a regular column / row';

      ws.pageSetup.printTitlesColumn = 'A:B';
      ws.pageSetup.printTitlesRow = '1:2';

      for (let i = 0; i < 100; i++) {
        ws.addRow(['repeated column, not repeated row', 'repeated column, not repeated row', 'no repeat', 'no repeat']);
      }

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('printHeaderAndColumn');
          expect(ws2.pageSetup.printTitlesRow).to.equal('1:2');
          expect(ws2.pageSetup.printTitlesColumn).to.equal('A:B');
        });
    });

    it('shared formula', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.fillFormula('A1:B2', 'ROW()+COLUMN()', [
        [2, 3],
        [3, 4],
      ]);
      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('Hello');
          expect(ws2.getCell('A1').value).to.deep.equal({
            formula: 'ROW()+COLUMN()',
            shareType: 'shared',
            ref: 'A1:B2',
            result: 2,
          });
          expect(ws2.getCell('B1').value).to.deep.equal({
            sharedFormula: 'A1',
            result: 3,
          });
          expect(ws2.getCell('A2').value).to.deep.equal({
            sharedFormula: 'A1',
            result: 3,
          });
          expect(ws2.getCell('B2').value).to.deep.equal({
            sharedFormula: 'A1',
            result: 4,
          });
        });
    });

    it('auto filter', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 1;
      ws.getCell('B1').value = 1;
      ws.getCell('A2').value = 2;
      ws.getCell('B2').value = 2;
      ws.getCell('A3').value = 3;
      ws.getCell('B3').value = 3;

      ws.autoFilter = 'A1:B1';

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('Hello');
          expect(ws2.autoFilter).to.equal('A1:B1');
        });
    });

    it('company, manager, etc', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'World!';
      wb.company = 'Cyber Sapiens, Ltd';
      wb.manager = 'Guyon Roche';
      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          expect(wb2.company).to.equal(wb.company);
          expect(wb2.manager).to.equal(wb.manager);
        });
    });

    it('title, subject, etc', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'World!';
      wb.title = 'the title';
      wb.subject = 'the subject';
      wb.keywords = 'the keywords';
      wb.category = 'the category';
      wb.description = 'the description';
      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          expect(wb2.title).to.equal(wb.title);
          expect(wb2.subject).to.equal(wb.subject);
          expect(wb2.keywords).to.equal(wb.keywords);
          expect(wb2.category).to.equal(wb.category);
          expect(wb2.description).to.equal(wb.description);
        });
    });

    it('language, revision and contentStatus', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'World!';
      wb.language = 'Klingon';
      wb.revision = 2;
      wb.contentStauts = 'Final';
      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          expect(wb2.language).to.equal(wb.language);
          expect(wb2.revision).to.equal(wb.revision);
          expect(wb2.contentStatus).to.equal(wb.contentStatus);
        });
    });

    it('empty strings', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'Foo';
      ws.getCell('A2').value = '';
      ws.getCell('A3').value = 'Baz';
      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('Hello');

          expect(ws2.getCell('A1').value).to.equal('Foo');
          expect(ws2.getCell('A2').value).to.equal('');
          expect(ws2.getCell('A3').value).to.equal('Baz');
        });
    });

    it('dataValidations', () => {
      const wb = testUtils.createTestBook(new ExcelJS.Workbook(), 'xlsx', ['dataValidations']);

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          testUtils.checkTestBook(wb2, 'xlsx', ['dataValidations']);
        });
    });

    it('empty string', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet();

      ws.columns = [
        { key: 'id', width: 10 },
        { key: 'name', width: 32 },
      ];

      ws.addRow({ id: 1, name: '' });

      return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
    });

    it('a lot of sheets to xlsx file', () => {
      let i;
      const wb = new ExcelJS.Workbook();
      const numSheets = 90;
      // add numSheets sheets
      for (i = 1; i <= numSheets; i++) {
        const ws = wb.addWorksheet(`sheet${i}`);
        ws.getCell('A1').value = i;
      }
      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          for (i = 1; i <= numSheets; i++) {
            const ws2 = wb2.getWorksheet(`sheet${i}`);
            expect(ws2).to.be.ok;
            expect(ws2.getCell('A1').value).to.equal(i);
          }
        });
    }, 10000);

    it('csv file', () => {
      const wb = testUtils.createTestBook(new ExcelJS.Workbook(), 'csv');

      return wb.csv
        .writeFile(TEST_CSV_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.csv.readFile(TEST_CSV_FILE_NAME).then(() => wb2);
        })
        .then((wb2) => {
          testUtils.checkTestBook(wb2, 'csv');
        });
    }, 5000);

    it('CSV file and its configuration', () => {
      const writeOptions = {
        dateFormat: 'DD/MM/YYYY HH:mm:ss',
        dateUTC: false,
        encoding: 'utf-8',
        includeEmptyRows: false,
        sheetName: 'sheet1',
        formatterOptions: {
          delimiter: '\t',
          quote: false,
        },
      };
      const readOptions = {
        dateFormats: ['DD/MM/YYYY HH:mm:ss'],
        sheetName: 'sheet1',
        parserOptions: {
          delimiter: '\t',
          quote: false,
        },
      };
      const wb = testUtils.createTestBook(new ExcelJS.Workbook(), 'csv');

      return wb.csv
        .writeFile(TEST_CSV_FILE_NAME, writeOptions)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.csv.readFile(TEST_CSV_FILE_NAME, readOptions).then(() => wb2);
        })
        .then((wb2) => {
          testUtils.checkTestBook(wb2, 'csv', false, writeOptions);
        });
    }, 5000);

    it('defined names', () => {
      const wb1 = new ExcelJS.Workbook();
      const ws1a = wb1.addWorksheet('blort');
      const ws1b = wb1.addWorksheet('foo');

      function assign(sheet, address, value, name) {
        const cell = sheet.getCell(address);
        cell.value = value;
        if (Array.isArray(name)) {
          cell.names = name;
        } else {
          cell.name = name;
        }
      }

      // single entry
      assign(ws1a, 'A1', 5, 'five');

      // three amigos - horizontal line
      assign(ws1a, 'A3', 3, 'amigos');
      assign(ws1a, 'B3', 3, 'amigos');
      assign(ws1a, 'C3', 3, 'amigos');

      // three amigos - vertical line
      assign(ws1a, 'E1', 3, 'verts');
      assign(ws1a, 'E2', 3, 'verts');
      assign(ws1a, 'E3', 3, 'verts');

      // four square
      assign(ws1a, 'C5', 4, 'squares');
      assign(ws1a, 'B6', 4, 'squares');
      assign(ws1a, 'C6', 4, 'squares');
      assign(ws1a, 'B5', 4, 'squares');

      // long distance
      assign(ws1a, 'B7', 2, 'sheets');
      assign(ws1b, 'B7', 2, 'sheets');

      // two names
      assign(ws1a, 'G1', 1, 'thing1');
      ws1a.getCell('G1').addName('thing2');

      // once removed
      assign(ws1a, 'G2', 1, ['once', 'twice']);
      ws1a.getCell('G2').removeName('once');

      return wb1.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2a = wb2.getWorksheet('blort');
          const ws2b = wb2.getWorksheet('foo');

          function check(sheet, address, value, name) {
            const cell = sheet.getCell(address);
            expect(cell.value).to.equal(value);
            expect(cell.name).to.equal(name);
          }

          // single entry
          check(ws2a, 'A1', 5, 'five');

          // three amigos - horizontal line
          check(ws2a, 'A3', 3, 'amigos');
          check(ws2a, 'B3', 3, 'amigos');
          check(ws2a, 'C3', 3, 'amigos');

          // three amigos - vertical line
          check(ws2a, 'E1', 3, 'verts');
          check(ws2a, 'E2', 3, 'verts');
          check(ws2a, 'E3', 3, 'verts');

          // four square
          check(ws2a, 'C5', 4, 'squares');
          check(ws2a, 'B6', 4, 'squares');
          check(ws2a, 'C6', 4, 'squares');
          check(ws2a, 'B5', 4, 'squares');

          // long distance
          check(ws2a, 'B7', 2, 'sheets');
          check(ws2b, 'B7', 2, 'sheets');

          // two names
          expect(ws2a.getCell('G1').names).to.have.members(['thing1', 'thing2']);

          // once removed
          expect(ws2a.getCell('G2').names).to.have.members(['twice']);

          // ranges
          function rangeCheck(name, members) {
            const ranges = wb2.definedNames.getRanges(name);
            expect(ranges.name).to.equal(name);
            if (members.length) {
              expect(ranges.ranges).to.have.members(members);
            } else {
              expect(ranges.ranges.length).to.equal(0);
            }
          }

          rangeCheck('five', ['blort!$A$1']);
          rangeCheck('amigos', ['blort!$A$3:$C$3']);
          rangeCheck('verts', ['blort!$E$1:$E$3']);
          rangeCheck('squares', ['blort!$B$5:$C$6']);
          rangeCheck('sheets', ['blort!$B$7', 'foo!$B$7']);
          rangeCheck('thing1', ['blort!$G$1']);
          rangeCheck('thing2', ['blort!$G$1']);
          rangeCheck('once', []);
          rangeCheck('twice', ['blort!$G$2']);
        });
    });

    describe('Duplicate Rows', () => {
      it('Duplicate rows with styles properly', () => {
        const fileDuplicateRowTestFile = './spec/integration/data/duplicateRowTest.xlsx';
        const wb = new ExcelJS.Workbook();
        return wb.xlsx.readFile(fileDuplicateRowTestFile).then(() => {
          const ws = wb.getWorksheet('duplicateTest');

          ws.getCell('A1').value = 'OneInfo';
          ws.getCell('A2').value = 'TwoInfo';
          ws.duplicateRow(1, 2);

          return wb.xlsx
            .writeFile(TEST_XLSX_FILE_NAME)
            .then(() => {
              const wb2 = new ExcelJS.Workbook();
              return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
            })
            .then((wb2) => {
              const ws2 = wb2.getWorksheet('duplicateTest');

              expect(ws2.getCell('A2').value).to.equal('OneInfo');
              expect(ws2.getCell('A2').style).to.equal(ws2.getCell('A1').style);
              expect(ws2.getCell('A3').value).to.equal('OneInfo');
              expect(ws2.getCell('A3').style).to.equal(ws2.getCell('A1').style);
              expect(ws2.getCell('A4').value).to.be.null;
            });
        });
      });

      it('Duplicate rows replacing properly', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('duplicateTest');
        ws.getCell('A1').value = 'OneInfo';
        ws.getCell('A2').value = 'TwoInfo';
        ws.getCell('A3').value = 'ThreeInfo';
        ws.getCell('A4').value = 'FourInfo';
        ws.duplicateRow(1, 2, false);

        return wb.xlsx
          .writeFile(TEST_XLSX_FILE_NAME)
          .then(() => {
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
          })
          .then((wb2) => {
            const ws2 = wb2.getWorksheet('duplicateTest');

            expect(ws2.getCell('A1').value).to.equal('OneInfo');
            expect(ws2.getCell('A2').value).to.equal('OneInfo');
            expect(ws2.getCell('A3').value).to.equal('OneInfo');
            expect(ws2.getCell('A4').value).to.equal('FourInfo');
          });
      });

      it('Duplicate rows shifting properly', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('duplicateTest');
        ws.getCell('A1').value = 'OneInfo';
        ws.getCell('A2').value = 'TwoInfo';
        ws.getCell('A3').value = 'ThreeInfo';
        ws.getCell('A4').value = 'FourInfo';
        ws.duplicateRow(1, 2, true);

        return wb.xlsx
          .writeFile(TEST_XLSX_FILE_NAME)
          .then(() => {
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
          })
          .then((wb2) => {
            const ws2 = wb2.getWorksheet('duplicateTest');

            expect(ws2.getCell('A1').value).to.equal('OneInfo');
            expect(ws2.getCell('A2').value).to.equal('OneInfo');
            expect(ws2.getCell('A3').value).to.equal('OneInfo');
            expect(ws2.getCell('A4').value).to.equal('TwoInfo');
          });
      });

      it('Duplicate rows with height properly', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('duplicateTest');
        ws.getCell('A1').value = 'OneInfo';
        ws.getCell('A2').value = 'TwoInfo';
        ws.getRow(1).height = 25;
        ws.getRow(2).height = 15;
        ws.duplicateRow(1, 1, true);

        return wb.xlsx
          .writeFile(TEST_XLSX_FILE_NAME)
          .then(() => {
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
          })
          .then((wb2) => {
            const ws2 = wb2.getWorksheet('duplicateTest');

            expect(ws2.getCell('A1').value).to.equal('OneInfo');
            expect(ws2.getCell('A2').value).to.equal('OneInfo');
            expect(ws2.getRow(1).height).to.equal(ws2.getRow(2).height);
            expect(ws2.getRow(1).height).to.not.equal(ws2.getRow(3).height);
          });
      });
    });

    describe('Merge Cells', () => {
      it('serialises and deserialises properly', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');

        // initial values
        ws.getCell('B2').value = 'B2';

        ws.mergeCells('B2:C3');

        return wb.xlsx
          .writeFile(TEST_XLSX_FILE_NAME)
          .then(() => {
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
          })
          .then((wb2) => {
            const ws2 = wb2.getWorksheet('blort');

            expect(ws2.getCell('B2').value).to.equal('B2');
            expect(ws2.getCell('B3').value).to.equal('B2');
            expect(ws2.getCell('C2').value).to.equal('B2');
            expect(ws2.getCell('C3').value).to.equal('B2');

            expect(ws2.getCell('B2').type).to.equal(ExcelJS.ValueType.String);
            expect(ws2.getCell('B3').type).to.equal(ExcelJS.ValueType.Merge);
            expect(ws2.getCell('C2').type).to.equal(ExcelJS.ValueType.Merge);
            expect(ws2.getCell('C3').type).to.equal(ExcelJS.ValueType.Merge);
          });
      });

      it('styles', () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');

        // initial values
        const B2 = ws.getCell('B2');
        B2.value = 5;
        B2.style.font = testUtils.styles.fonts.broadwayRedOutline20;
        B2.style.border = testUtils.styles.borders.doubleRed;
        B2.style.fill = testUtils.styles.fills.blueWhiteHGrad;
        B2.style.alignment = testUtils.styles.namedAlignments.middleCentre;
        B2.style.numFmt = testUtils.styles.numFmts.numFmt1;

        // expecting styles to be copied (see worksheet spec)
        ws.mergeCells('B2:C3');

        return wb.xlsx
          .writeFile(TEST_XLSX_FILE_NAME)
          .then(() => {
            const wb2 = new ExcelJS.Workbook();
            return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
          })
          .then((wb2) => {
            const ws2 = wb2.getWorksheet('blort');

            expect(ws2.getCell('B2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
            expect(ws2.getCell('B2').border).to.deep.equal(testUtils.styles.borders.doubleRed);
            expect(ws2.getCell('B2').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
            expect(ws2.getCell('B2').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
            expect(ws2.getCell('B2').numFmt).to.equal(testUtils.styles.numFmts.numFmt1);

            expect(ws2.getCell('B3').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
            expect(ws2.getCell('B3').border).to.deep.equal(testUtils.styles.borders.doubleRed);
            expect(ws2.getCell('B3').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
            expect(ws2.getCell('B3').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
            expect(ws2.getCell('B3').numFmt).to.equal(testUtils.styles.numFmts.numFmt1);

            expect(ws2.getCell('C2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
            expect(ws2.getCell('C2').border).to.deep.equal(testUtils.styles.borders.doubleRed);
            expect(ws2.getCell('C2').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
            expect(ws2.getCell('C2').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
            expect(ws2.getCell('C2').numFmt).to.equal(testUtils.styles.numFmts.numFmt1);

            expect(ws2.getCell('C3').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
            expect(ws2.getCell('C3').border).to.deep.equal(testUtils.styles.borders.doubleRed);
            expect(ws2.getCell('C3').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
            expect(ws2.getCell('C3').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
            expect(ws2.getCell('C3').numFmt).to.equal(testUtils.styles.numFmts.numFmt1);
          });
      });
    });
  });

  it('spliced meat and ham', () => {
    const wb = new ExcelJS.Workbook();
    const sheets = [
      'splice.rows.removeOnly',
      'splice.rows.insertFewer',
      'splice.rows.insertSame',
      'splice.rows.insertMore',
      'splice.rows.insertStyle',
      'splice.columns.removeOnly',
      'splice.columns.insertFewer',
      'splice.columns.insertSame',
      'splice.columns.insertMore',
    ];
    const options = {
      checkBadAlignments: false,
      checkSheetProperties: false,
      checkViews: false,
    };

    testUtils.createTestBook(wb, 'xlsx', sheets, options);

    return wb.xlsx
      .writeFile(TEST_XLSX_FILE_NAME)
      .then(() => {
        const wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      })
      .then((wb2) => {
        testUtils.checkTestBook(wb2, 'xlsx', sheets, options);
      });
  });

  it('throws an error when xlsx file not found', () => {
    const wb = new ExcelJS.Workbook();
    let success = 0;
    let failure;
    return wb.xlsx
      .readFile('./wb.doesnotexist.xlsx')
      .then((/* wb2 */) => {
        success = 1;
      })
      .catch((error) => {
        success = 2;
        failure = error;
      })
      .then(() => {
        expect(success).to.equal(2);
        expect(failure).to.be.an.instanceOf(Error);
        expect(failure.message).to.equal('File not found: ./wb.doesnotexist.xlsx');
      });
  });

  it('throws an error when csv file not found', () => {
    const wb = new ExcelJS.Workbook();
    let success = 0;
    let failure;
    return wb.csv
      .readFile('./wb.doesnotexist.csv')
      .then((/* wb */) => {
        success = 1;
      })
      .catch((error) => {
        success = 2;
        failure = error;
      })
      .then(() => {
        expect(success).to.equal(2);
        expect(failure).to.be.an.instanceOf(Error);
        expect(failure.message).to.equal('File not found: ./wb.doesnotexist.csv');
      });
  });
  it('throw an error for wrong data type', async () => {
    const wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load({});
      expect.fail('should fail for given argument');
    } catch (e) {
      expect(e.message).to.equal(
        "Can't read the data of 'the loaded zip file'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?",
      );
    }
  });

  describe('Xlsx container options', () => {
    const buildBook = () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('opts');
      ws.getCell('A1').value = 'plain';
      ws.getCell('A1').font = { bold: true };
      ws.getCell('B1').value = 42;
      return wb;
    };

    it('loads base64 encoded data', async () => {
      const buffer = await buildBook().xlsx.writeBuffer();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer.toString('base64'), { base64: true });

      const ws2 = wb2.getWorksheet('opts');
      expect(ws2.getCell('A1').value).to.equal('plain');
      expect(ws2.getCell('B1').value).to.equal(42);
    });

    it('createInputStream is deprecated', () => {
      const wb = new ExcelJS.Workbook();
      expect(() => wb.xlsx.createInputStream()).to.throw('`XLSX#createInputStream` is deprecated');
    });

    it('writeFile rejects when the target cannot be created', async () => {
      let error;
      try {
        await buildBook().xlsx.writeFile('./spec/out/no-such-dir-xlsx-container/wb.test.xlsx');
      } catch (e) {
        error = e;
      }
      expect(error).to.be.an('error');
      expect(error.code).to.equal('ENOENT');
    });

    it('useStyles:false drops style information but keeps values', async () => {
      const buffer = await buildBook().xlsx.writeBuffer({ useStyles: false });

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet('opts');
      expect(ws2.getCell('A1').value).to.equal('plain');
      expect(ws2.getCell('B1').value).to.equal(42);
      expect(ws2.getCell('A1').font).to.be.undefined;
    });

    it('keeps styles by default', async () => {
      const buffer = await buildBook().xlsx.writeBuffer();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);

      expect(wb2.getWorksheet('opts').getCell('A1').font).to.deep.equal({ bold: true });
    });

    // cell-xform decodes t="str" text (and string formula results) that the XML parser already decoded
    describe('literal entity text', () => {
      const roundTrip = async (value) => {
        const wb = new ExcelJS.Workbook();
        wb.addWorksheet('entities').getCell('A1').value = value;
        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(await wb.xlsx.writeBuffer({ useSharedStrings: false }));
        return wb2.getWorksheet('entities').getCell('A1').value;
      };

      knownBug(
        'keeps a string containing an entity reference',
        async () => {
          // the streaming writer emits plain strings as t="str" cells
          const filename = './spec/out/wb-entity-str.test.xlsx';
          const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
          const ws = writer.addWorksheet('entities');
          ws.addRow(['Tom &amp; Jerry']).commit();
          ws.commit();
          await writer.commit();

          const wb2 = new ExcelJS.Workbook();
          await wb2.xlsx.readFile(filename);
          expect(wb2.getWorksheet('entities').getCell('A1').value).to.equal('Tom &amp; Jerry');
        },
        /expected 'Tom & Jerry' to equal 'Tom &amp; Jerry'/,
      );

      knownBug(
        'keeps a string formula result containing an entity reference',
        async () => {
          const value = await roundTrip({ formula: '"R &amp; D"', result: 'R &amp; D' });
          expect(value.result).to.equal('R &amp; D');
        },
        /expected 'R & D' to equal 'R &amp; D'/,
      );
    });
  });

  describe('Sheet Views', () => {
    it('frozen panes', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('frozen');
      ws.views = [
        {
          state: 'frozen',
          xSplit: 2,
          ySplit: 3,
          topLeftCell: 'C4',
          activeCell: 'D5',
        },
        { state: 'frozen', ySplit: 1 },
        { state: 'frozen', xSplit: 1 },
      ];
      ws.getCell('A1').value = 'Let it Snow!';

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('frozen');
          expect(ws2).to.be.ok;
          expect(ws2.getCell('A1').value).to.equal('Let it Snow!');
          expect(ws2.views).to.deep.equal([
            {
              workbookViewId: 0,
              state: 'frozen',
              xSplit: 2,
              ySplit: 3,
              topLeftCell: 'C4',
              activeCell: 'D5',
              showRuler: true,
              showGridLines: true,
              showRowColHeaders: true,
              zoomScale: 100,
              zoomScaleNormal: 100,
              rightToLeft: false,
            },
            {
              workbookViewId: 0,
              state: 'frozen',
              xSplit: 0,
              ySplit: 1,
              topLeftCell: 'A2',
              showRuler: true,
              showGridLines: true,
              showRowColHeaders: true,
              zoomScale: 100,
              zoomScaleNormal: 100,
              rightToLeft: false,
            },
            {
              workbookViewId: 0,
              state: 'frozen',
              xSplit: 1,
              ySplit: 0,
              topLeftCell: 'B1',
              showRuler: true,
              showGridLines: true,
              showRowColHeaders: true,
              zoomScale: 100,
              zoomScaleNormal: 100,
              rightToLeft: false,
            },
          ]);
        });
    });

    it('serialises split panes', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('split');
      ws.views = [
        {
          state: 'split',
          xSplit: 2000,
          ySplit: 3000,
          topLeftCell: 'C4',
          activeCell: 'D5',
          activePane: 'bottomRight',
        },
        {
          state: 'split',
          ySplit: 1500,
          activePane: 'bottomLeft',
          topLeftCell: 'A10',
        },
        { state: 'split', xSplit: 1500, activePane: 'topRight' },
      ];
      ws.getCell('A1').value = 'Do the splits!';

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('split');
          expect(ws2).to.be.ok;
          expect(ws2.getCell('A1').value).to.equal('Do the splits!');
          expect(ws2.views).to.deep.equal([
            {
              workbookViewId: 0,
              state: 'split',
              xSplit: 2000,
              ySplit: 3000,
              topLeftCell: 'C4',
              activeCell: 'D5',
              activePane: 'bottomRight',
              showRuler: true,
              showGridLines: true,
              showRowColHeaders: true,
              zoomScale: 100,
              zoomScaleNormal: 100,
              rightToLeft: false,
            },
            {
              workbookViewId: 0,
              state: 'split',
              xSplit: 0,
              ySplit: 1500,
              topLeftCell: 'A10',
              activePane: 'bottomLeft',
              showRuler: true,
              showGridLines: true,
              showRowColHeaders: true,
              zoomScale: 100,
              zoomScaleNormal: 100,
              rightToLeft: false,
            },
            {
              workbookViewId: 0,
              state: 'split',
              xSplit: 1500,
              ySplit: 0,
              topLeftCell: undefined,
              activePane: 'topRight',
              showRuler: true,
              showGridLines: true,
              showRowColHeaders: true,
              zoomScale: 100,
              zoomScaleNormal: 100,
              rightToLeft: false,
            },
          ]);
        });
    });

    it('multiple book views', () => {
      const wb = new ExcelJS.Workbook();
      wb.views = [testUtils.views.book.visible, testUtils.views.book.hidden];

      const ws1 = wb.addWorksheet('one');
      ws1.views = [testUtils.views.sheet.frozen];

      const ws2 = wb.addWorksheet('two');
      ws2.views = [testUtils.views.sheet.split];

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          expect(wb2.views).to.deep.equal(wb.views);

          const ws1b = wb2.getWorksheet('one');
          expect(ws1b.views).to.deep.equal(ws1.views);

          const ws2b = wb2.getWorksheet('two');
          expect(ws2b.views).to.deep.equal(ws2.views);
        });
    });
  });

  describe('CSV', () => {
    function workbookWithRows(rows, name = 'data') {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet(name);
      rows.forEach((values, i) => {
        if (values) ws.getRow(i + 1).values = values;
      });
      return { wb, ws };
    }

    const bufferText = async (wb, options) => (await wb.csv.writeBuffer(options)).toString();

    async function readCsv(text, options) {
      const wb = new ExcelJS.Workbook();
      const ws = await wb.csv.read(Readable.from([text]), options);
      return { wb, ws };
    }

    // cell by cell, so that empty cells read as null instead of holes
    const rowValues = (ws, rowNumber = 1) => {
      const values = [];
      for (let col = 1; col <= ws.columnCount; col++) {
        values.push(ws.getCell(rowNumber, col).value);
      }
      return values;
    };

    describe('writeBuffer', () => {
      it('writes each row as a line without a trailing newline', async () => {
        const { wb } = workbookWithRows([
          ['a', 1],
          ['b', 2],
          ['c', 3],
        ]);
        expect(await bufferText(wb)).to.equal('a,1\nb,2\nc,3');
      });

      it('writes a blank line for a missing row by default', async () => {
        const { wb } = workbookWithRows([['a'], null, ['b']]);
        expect(await bufferText(wb)).to.equal('a\n\nb');
      });

      it('skips missing rows when includeEmptyRows is false', async () => {
        const { wb } = workbookWithRows([['a'], null, ['b']]);
        expect(await bufferText(wb, { includeEmptyRows: false })).to.equal('a\nb');
      });

      it('writes the first worksheet when no sheet is named', async () => {
        const { wb } = workbookWithRows([['first']], 'one');
        wb.addWorksheet('two').addRow(['second']);
        expect(await bufferText(wb)).to.equal('first');
      });

      it('selects the worksheet by sheetId or sheetName', async () => {
        const { wb } = workbookWithRows([['first']], 'one');
        wb.addWorksheet('two').addRow(['second']);
        expect(await bufferText(wb, { sheetId: 2 })).to.equal('second');
        expect(await bufferText(wb, { sheetName: 'two' })).to.equal('second');
        expect(await bufferText(wb, { sheetName: 'one' })).to.equal('first');
      });

      it('writes nothing for an unknown sheet', async () => {
        const { wb } = workbookWithRows([['first']]);
        expect(await bufferText(wb, { sheetName: 'nope' })).to.equal('');
      });

      it('writes dates in UTC ISO format when dateUTC is set', async () => {
        const { wb } = workbookWithRows([[new Date('2020-01-02T03:04:05Z')]]);
        expect(await bufferText(wb, { dateUTC: true })).to.equal('2020-01-02T03:04:05Z');
      });

      it('applies dateFormat with and without dateUTC', async () => {
        const { wb } = workbookWithRows([[new Date('2020-06-15T12:00:00Z')]]);
        expect(await bufferText(wb, { dateUTC: true, dateFormat: 'YYYY/MM/DD HH:mm' })).to.equal('2020/06/15 12:00');
        expect(await bufferText(wb, { dateFormat: 'YYYY' })).to.equal('2020');
      });

      it('writes text, hyperlink, error, boolean and formula values as plain fields', async () => {
        const { wb } = workbookWithRows([
          [
            'str',
            0,
            true,
            false,
            { text: 'label', hyperlink: 'http://example.com' },
            { error: '#N/A' },
            { formula: 'A1', result: 5 },
            { formula: 'A1', result: 'text' },
          ],
        ]);
        expect(await bufferText(wb)).to.equal('str,0,true,false,http://example.com,#N/A,5,text');
      });

      it('hands raw cell values to a custom map', async () => {
        const { wb } = workbookWithRows([['a', 2, { formula: 'A1', result: 'x' }]]);
        const seen = [];
        const text = await bufferText(wb, {
          map: (value) => {
            seen.push(value);
            return `<${typeof value}>`;
          },
        });
        expect(seen).to.deep.equal(['a', 2, { formula: 'A1', result: 'x' }]);
        expect(text).to.equal('<string>,<number>,<object>');
      });

      // known bug: falsy formula results (0, false) are written as empty fields because of `value.result || ''` — see spec/KNOWN-BUGS.md
      knownBug(
        'writes falsy formula results',
        async () => {
          const { wb } = workbookWithRows([
            [
              { formula: '1-1', result: 0 },
              { formula: 'A1', result: false },
            ],
          ]);
          expect(await bufferText(wb)).to.equal('0,false');
        },
        /expected ',' to equal '0,false'/,
      );

      // known bug (needs owner confirmation, may be by design): rich text cells are written as a JSON blob instead of their text — see spec/KNOWN-BUGS.md
      knownBug(
        'writes a rich text cell as its plain concatenated text',
        async () => {
          const { wb } = workbookWithRows([[{ richText: [{ text: 'Hel' }, { text: 'lo' }] }]]);
          expect(await bufferText(wb)).to.equal('Hello');
        },
        /richText.*to equal 'Hello'/s,
      );
    });

    describe('write', () => {
      it('writes to a stream and resolves once it has finished', async () => {
        const { wb } = workbookWithRows([
          ['a', 'b'],
          ['c', 'd'],
        ]);
        const stream = new PassThrough();
        const chunks = [];
        stream.on('data', (chunk) => chunks.push(chunk));
        await wb.csv.write(stream, { sheetName: 'data' });
        expect(Buffer.concat(chunks).toString()).to.equal('a,b\nc,d');
      });

      it('quotes fields that contain the delimiter, quotes or newlines', async () => {
        const { wb } = workbookWithRows([['a,b', 'say "hi"', 'two\nlines']]);
        expect(await bufferText(wb)).to.equal('"a,b","say ""hi""","two\nlines"');
      });
    });

    describe('read', () => {
      it('converts special values and numbers and keeps other text', async () => {
        const { ws } = await readCsv('x,#N/A,true,Infinity,1e3,,2020-01-02\n');
        expect(rowValues(ws)).to.deep.equal([
          'x',
          { error: '#N/A' },
          true,
          'Infinity',
          1000,
          null,
          new Date(2020, 0, 2),
        ]);
      });

      it('recognises every error literal and boolean spelling in lower case only', async () => {
        const { ws } = await readCsv('#REF!,#NAME?,#DIV/0!,#NULL!,#VALUE!,#NUM!,false,TRUE,True\n');
        expect(rowValues(ws)).to.deep.equal([
          { error: '#REF!' },
          { error: '#NAME?' },
          { error: '#DIV/0!' },
          { error: '#NULL!' },
          { error: '#VALUE!' },
          { error: '#NUM!' },
          false,
          'TRUE',
          'True',
        ]);
      });

      it('reads hexadecimal and exponent notation as numbers', async () => {
        const { ws } = await readCsv('0x10,1e3,-2.5,007\n');
        expect(rowValues(ws)).to.deep.equal([16, 1000, -2.5, 7]);
      });

      it('parses MM-DD-YYYY, YYYY-MM-DD and YYYY-MM-DDTHH:mm:ss as local dates', async () => {
        const { ws } = await readCsv('01-02-2020,2020-01-02,2020-01-02T03:04:05\n');
        expect(rowValues(ws)).to.deep.equal([
          new Date(2020, 0, 2),
          new Date(2020, 0, 2),
          new Date(2020, 0, 2, 3, 4, 5),
        ]);
      });

      it('leaves text that merely looks like a date as a string', async () => {
        const { ws } = await readCsv('2020-13-45,01/02/2020\n');
        expect(rowValues(ws)).to.deep.equal(['2020-13-45', '01/02/2020']);
      });

      // known bug: dateFormats 'YYYY-MM-DD[T]HH:mm:ssZ' is matched strictly against the local offset, so a UTC "...Z" timestamp (what write emits with dateUTC) stays a string — see spec/KNOWN-BUGS.md
      knownBug(
        'reads back the UTC timestamps that write emits with dateUTC',
        async () => {
          const original = new Date('2020-01-02T03:04:05Z');
          const { wb } = workbookWithRows([[original]]);
          const text = await bufferText(wb, { dateUTC: true });
          const { ws } = await readCsv(text);
          expect(ws.getCell('A1').value).to.deep.equal(original);
        },
        /expected '2020-01-02T03:04:05Z' to deeply equal 2020-01-02T03:04:05\.000Z/,
      );

      it('uses custom dateFormats', async () => {
        const { ws } = await readCsv('31/12/2020,2020-12-31\n', { dateFormats: ['DD/MM/YYYY'] });
        expect(rowValues(ws)).to.deep.equal([new Date(2020, 11, 31), '2020-12-31']);
      });

      it('lets options.map replace the default conversion', async () => {
        const { ws } = await readCsv('a,1,true\n', { map: (datum) => datum.toUpperCase() });
        expect(rowValues(ws)).to.deep.equal(['A', '1', 'TRUE']);
      });

      it('names the worksheet sheet1 by default and by options.sheetName when given', async () => {
        const defaults = await readCsv('a\n');
        expect(defaults.ws.name).to.equal('sheet1');
        expect(defaults.wb.worksheets).to.have.length(1);

        const named = await readCsv('a\n', { sheetName: 'imported' });
        expect(named.ws.name).to.equal('imported');
        expect(named.wb.getWorksheet('imported')).to.equal(named.ws);
      });

      it('reads multiple rows and keeps empty fields as null', async () => {
        const { ws } = await readCsv('a,b\n,d\n');
        expect(ws.rowCount).to.equal(2);
        expect(rowValues(ws, 2)).to.deep.equal([null, 'd']);
      });

      // known bug: parserOptions.headers makes fast-csv emit objects, but read() calls data.map on each row, so every read rejects — see spec/KNOWN-BUGS.md
      knownBug(
        'reads with parserOptions.headers',
        async () => {
          const { ws } = await readCsv('a,b\n1,2\n', { parserOptions: { headers: true } });
          expect(ws.rowCount).to.be.above(0);
        },
        /data\.map is not a function/,
      );

      it('passes parserOptions to the parser', async () => {
        const { ws } = await readCsv('a;b\n', { parserOptions: { delimiter: ';' } });
        expect(rowValues(ws)).to.deep.equal(['a', 'b']);
      });

      it('rejects when the parser fails on malformed input', async () => {
        const wb = new ExcelJS.Workbook();
        let failure;
        try {
          await wb.csv.read(Readable.from(['a,"b\n']));
        } catch (error) {
          failure = error;
        }
        expect(failure).to.be.an.instanceOf(Error);
        expect(failure.message).to.match(/missing closing/);
      });

      it('rejects when the parser sees a stray character after a closing quote', async () => {
        const wb = new ExcelJS.Workbook();
        let failure;
        try {
          await wb.csv.read(Readable.from(['a,b\n"c"x,d\n']));
        } catch (error) {
          failure = error;
        }
        expect(failure).to.be.an.instanceOf(Error);
        expect(failure.message).to.match(/expected: ','/);
      });
    });

    describe('createInputStream', () => {
      it('is deprecated and throws', () => {
        const wb = new ExcelJS.Workbook();
        expect(() => wb.csv.createInputStream()).to.throw(/deprecated/);
        expect(() => wb.csv.createInputStream()).to.throw(/CSV#read/);
      });
    });
  });
});
