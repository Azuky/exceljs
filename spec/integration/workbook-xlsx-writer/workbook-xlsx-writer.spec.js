const fs = require('fs');
const { promisify } = require('util');
const { Writable } = require('stream');

const testUtils = require('../../utils/index');
const { knownBug } = require('../../utils/known-bug');

const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';
const FILL_FORMULA_FILE = './spec/out/wb-fill-formula.test.xlsx';
const ADD_WORKSHEET_FILE = './spec/out/wb-add-worksheet.test.xlsx';
const TAB_COLOR_FILE = './spec/out/wb-tab-color.test.xlsx';
const IMAGE_SOURCE_FILE = './spec/out/wb-image-source.test.xlsx';
const IMAGE_FILENAME = `${__dirname}/../data/image.png`;
const fsReadFileAsync = promisify(fs.readFile);

describe('WorkbookWriter', () => {
  it('creates sheets with correct names', () => {
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter();
    const ws1 = wb.addWorksheet('Hello, World!');
    expect(ws1.name).to.equal('Hello, World!');

    const ws2 = wb.addWorksheet();
    expect(ws2.name).to.match(/sheet\d+/);
  });

  describe('Serialise', () => {
    it('xlsx file', () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true,
      };
      const wb = testUtils.createTestBook(new ExcelJS.stream.xlsx.WorkbookWriter(options), 'xlsx');

      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          testUtils.checkTestBook(wb2, 'xlsx', undefined, { checkFillFormula: false });
        });
    });

    it('shared formula', () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: false,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = {
        formula: 'ROW()+COLUMN()',
        ref: 'A1:B2',
        result: 2,
      };
      ws.getCell('B1').value = { sharedFormula: 'A1', result: 3 };
      ws.getCell('A2').value = { sharedFormula: 'A1', result: 3 };
      ws.getCell('B2').value = { sharedFormula: 'A1', result: 4 };

      ws.commit();
      return wb
        .commit()
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

    // WorksheetWriter has no fillFormula, so the shared checkSheet() skips the A10:E11 block for
    // the stream writer (it only runs when the sheet carries those cells). Assert it explicitly.
    it('fill-style shared formula and array formula', async () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: FILL_FORMULA_FILE });
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A9').value = 1;
      ws.getCell('A10').value = { shareType: 'shared', ref: 'A10:E10', formula: 'A9', result: 1 };
      ['B10', 'C10', 'D10', 'E10'].forEach((address, index) => {
        ws.getCell(address).value = { sharedFormula: 'A10', result: index + 2 };
      });
      ws.getCell('A11').value = { shareType: 'array', ref: 'A11:E11', formula: 'A9', result: 1 };
      ['B11', 'C11', 'D11', 'E11'].forEach((address) => {
        ws.getCell(address).value = 1;
      });
      ws.commit();
      await wb.commit();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(FILL_FORMULA_FILE);
      const ws2 = wb2.getWorksheet('Hello');

      expect(ws2.getCell('A10').value).to.deep.equal({
        shareType: 'shared',
        ref: 'A10:E10',
        formula: 'A9',
        result: 1,
      });
      ['B10', 'C10', 'D10', 'E10'].forEach((address, index) => {
        expect(ws2.getCell(address).value).to.deep.equal({
          sharedFormula: 'A10',
          result: index + 2,
        });
        expect(ws2.getCell(address).formula).to.equal(`${address[0]}9`);
      });
      expect(ws2.getCell('A11').value).to.deep.equal({
        shareType: 'array',
        ref: 'A11:E11',
        formula: 'A9',
        result: 1,
      });
      ['B11', 'C11', 'D11', 'E11'].forEach((address) => {
        expect(ws2.getCell(address).value).to.equal(1);
      });
    });

    it('auto filter', () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: false,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 1;
      ws.getCell('B1').value = 1;
      ws.getCell('A2').value = 2;
      ws.getCell('B2').value = 2;
      ws.getCell('A3').value = 3;
      ws.getCell('B3').value = 3;

      ws.autoFilter = 'A1:B1';
      ws.commit();

      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('Hello');
          expect(ws2.autoFilter).to.equal('A1:B1');
        });
    });

    describe('addWorksheet options', () => {
      it('keeps state, views, page setup, auto filter, header/footer and properties', async () => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: ADD_WORKSHEET_FILE });
        const visible = wb.addWorksheet('visible');
        visible.getCell('A1').value = 1;
        const hidden = wb.addWorksheet('hidden', {
          state: 'hidden',
          views: [{ state: 'frozen', xSplit: 1, ySplit: 2, topLeftCell: 'B3', activeCell: 'B3' }],
          pageSetup: { orientation: 'landscape' },
          autoFilter: 'A1:B1',
          headerFooter: { oddHeader: 'Header', oddFooter: 'Footer' },
          properties: { tabColor: { argb: 'FF00FF00' }, defaultColWidth: 20 },
        });
        hidden.getCell('A1').value = 'x';
        hidden.getCell('B1').value = 'y';
        visible.commit();
        hidden.commit();
        await wb.commit();

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.readFile(ADD_WORKSHEET_FILE);

        const visible2 = wb2.getWorksheet('visible');
        expect(visible2.state).to.equal('visible');
        expect(visible2.pageSetup.orientation).to.equal('portrait');

        const hidden2 = wb2.getWorksheet('hidden');
        expect(hidden2.state).to.equal('hidden');
        expect(hidden2.views).to.have.length(1);
        expect(hidden2.views[0]).to.deep.include({
          state: 'frozen',
          xSplit: 1,
          ySplit: 2,
          topLeftCell: 'B3',
          activeCell: 'B3',
        });
        expect(hidden2.pageSetup.orientation).to.equal('landscape');
        expect(hidden2.autoFilter).to.equal('A1:B1');
        expect(hidden2.headerFooter.oddHeader).to.equal('Header');
        expect(hidden2.headerFooter.oddFooter).to.equal('Footer');
        expect(hidden2.properties.tabColor).to.deep.equal({ argb: 'FF00FF00' });
        expect(hidden2.properties.defaultColWidth).to.equal(20);
      });

      it('moves the deprecated tabColor option into properties and warns', async () => {
        const trace = vi.spyOn(console, 'trace').mockImplementation(() => {});
        try {
          const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: TAB_COLOR_FILE });
          const ws = wb.addWorksheet('colored', { tabColor: { argb: 'FFFF0000' } });
          ws.getCell('A1').value = 1;
          ws.commit();
          await wb.commit();

          expect(trace).toHaveBeenCalledTimes(1);
          expect(trace.mock.calls[0][0]).to.equal('tabColor option has moved to { properties: tabColor: {...} }');
        } finally {
          trace.mockRestore();
        }

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.readFile(TAB_COLOR_FILE);
        expect(wb2.getWorksheet('colored').properties.tabColor).to.deep.equal({ argb: 'FFFF0000' });
      });
    });

    it('Without styles', () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: false,
      };
      const wb = testUtils.createTestBook(new ExcelJS.stream.xlsx.WorkbookWriter(options), 'xlsx');

      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          testUtils.checkTestBook(wb2, 'xlsx', undefined, {
            checkFillFormula: false,
            checkStyles: false,
          });
        });
    });

    it('serializes row styles and columns properly', () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('blort');

      const colStyle = {
        font: testUtils.styles.fonts.comicSansUdB16,
        alignment: testUtils.styles.namedAlignments.middleCentre,
      };
      ws.columns = [
        { header: 'A1', width: 10 },
        { header: 'B1', style: colStyle },
        { header: 'C1', width: 30 },
        { header: 'D1' },
      ];

      ws.getRow(2).font = testUtils.styles.fonts.broadwayRedOutline20;

      ws.getCell('A2').value = 'A2';
      ws.getCell('B2').value = 'B2';
      ws.getCell('C2').value = 'C2';
      ws.getCell('A3').value = 'A3';
      ws.getCell('B3').value = 'B3';
      ws.getCell('C3').value = 'C3';

      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('blort');
          ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'].forEach((address) => {
            expect(ws2.getCell(address).value).to.equal(address);
          });
          expect(ws2.getCell('B1').font).to.deep.equal(testUtils.styles.fonts.comicSansUdB16);
          expect(ws2.getCell('B1').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
          expect(ws2.getCell('A2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell('B2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell('C2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell('B3').font).to.deep.equal(testUtils.styles.fonts.comicSansUdB16);
          expect(ws2.getCell('B3').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);

          expect(ws2.getColumn(2).font).to.deep.equal(testUtils.styles.fonts.comicSansUdB16);
          expect(ws2.getColumn(2).alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
          expect(ws2.getColumn(2).width).to.equal(9);

          expect(ws2.getColumn(4).width).to.equal(undefined);

          expect(ws2.getRow(2).font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
        });
    });

    it('rich text', () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('Hello');

      ws.getCell('A1').value = {
        richText: [
          {
            font: { color: { argb: 'FF0000' } },
            text: 'red ',
          },
          {
            font: { color: { argb: '00FF00' }, bold: true },
            text: ' bold green',
          },
        ],
      };

      ws.getCell('B1').value = 'plain text';

      ws.commit();
      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('Hello');
          expect(ws2.getCell('A1').value).to.deep.equal({
            richText: [
              {
                font: { color: { argb: 'FF0000' } },
                text: 'red ',
              },
              {
                font: { color: { argb: '00FF00' }, bold: true },
                text: ' bold green',
              },
            ],
          });
          expect(ws2.getCell('B1').value).to.equal('plain text');
        });
    });

    it('A lot of sheets', () => {
      let i;
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: TEST_XLSX_FILE_NAME,
      });
      const numSheets = 90;
      // add numSheets sheets
      for (i = 1; i <= numSheets; i++) {
        const ws = wb.addWorksheet(`sheet${i}`);
        ws.getCell('A1').value = i;
      }
      return wb
        .commit()
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
    }, 5000);

    it('addRow', async () => {
      const options = {
        stream: fs.createWriteStream(TEST_XLSX_FILE_NAME, { flags: 'w' }),
        useStyles: true,
        useSharedStrings: true,
      };
      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const worksheet = workbook.addWorksheet('test');
      const newRow = worksheet.addRow(['hello']);
      newRow.commit();
      worksheet.commit();
      await workbook.commit();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.getWorksheet('test').getCell('A1').value).to.equal('hello');
    });

    it('defined names', () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: TEST_XLSX_FILE_NAME,
      });
      const ws = wb.addWorksheet('blort');
      ws.getCell('A1').value = 5;
      ws.getCell('A1').name = 'five';

      ws.getCell('A3').value = 'drei';
      ws.getCell('A3').name = 'threes';
      ws.getCell('B3').value = 'trois';
      ws.getCell('B3').name = 'threes';
      ws.getCell('B3').value = 'san';
      ws.getCell('B3').name = 'threes';

      ws.getCell('E1').value = 'grün';
      ws.getCell('E1').name = 'greens';
      ws.getCell('E2').value = 'vert';
      ws.getCell('E2').name = 'greens';
      ws.getCell('E3').value = 'verde';
      ws.getCell('E3').name = 'greens';

      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('blort');
          expect(ws2.getCell('A1').name).to.equal('five');

          expect(ws2.getCell('A3').name).to.equal('threes');
          expect(ws2.getCell('B3').name).to.equal('threes');
          expect(ws2.getCell('B3').value).to.equal('san');

          expect(ws2.getCell('E1').name).to.equal('greens');
          expect(ws2.getCell('E2').name).to.equal('greens');
          expect(ws2.getCell('E3').name).to.equal('greens');
        });
    });

    it('does not escape special xml characters', () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: TEST_XLSX_FILE_NAME,
        useSharedStrings: true,
      });
      const ws = wb.addWorksheet('blort');
      const xmlCharacters = 'xml characters: & < > "';

      ws.getCell('A1').value = xmlCharacters;

      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          const ws2 = wb2.getWorksheet('blort');
          expect(ws2.getCell('A1').value).to.equal(xmlCharacters);
        });
    });

    it('serializes and deserializes dataValidations', () => {
      const options = { filename: TEST_XLSX_FILE_NAME };
      const wb = testUtils.createTestBook(new ExcelJS.stream.xlsx.WorkbookWriter(options), 'xlsx', ['dataValidations']);

      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          testUtils.checkTestBook(wb2, 'xlsx', ['dataValidations']);
        });
    });

    it('with zip compression option', () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true,
        zip: {
          zlib: { level: 9 }, // Sets the compression level.
        },
      };
      const wb = testUtils.createTestBook(new ExcelJS.stream.xlsx.WorkbookWriter(options), 'xlsx', ['dataValidations']);

      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          testUtils.checkTestBook(wb2, 'xlsx', ['dataValidations']);
        });
    });

    it('zip compression level changes the size of the output', async () => {
      const sizeAtLevel = async (level) => {
        const filename = `./spec/out/wb-zip-level-${level}.test.xlsx`;
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
          filename,
          useSharedStrings: false,
          zip: { zlib: { level } },
        });
        const ws = wb.addWorksheet('zip');
        for (let i = 1; i <= 200; i++) {
          ws.addRow(['the same repeated text', 'the same repeated text']).commit();
        }
        ws.commit();
        await wb.commit();
        return (await fs.promises.stat(filename)).size;
      };

      const stored = await sizeAtLevel(0);
      const compressed = await sizeAtLevel(9);
      expect(compressed).to.be.lessThan(stored / 2);
    });

    it('writes notes', async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('Hello');
      ws.getCell('B2').value = 5;
      ws.getCell('B2').note = 'five';

      const note = {
        texts: [
          {
            font: {
              size: 12,
              color: { argb: 'FFFF6600' },
              name: 'Calibri',
              scheme: 'minor',
            },
            text: 'seven',
          },
        ],
        margins: {
          insetmode: 'auto',
          inset: [0.13, 0.13, 0.25, 0.25],
        },
        protection: {
          locked: 'True',
          lockText: 'True',
        },
        editAs: 'twoCells',
      };
      ws.getCell('D2').value = 7;
      ws.getCell('D2').note = note;

      await wb.commit();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet('Hello');

      expect(ws2.getCell('B2').value).to.equal(5);
      expect(ws2.getCell('B2').note).to.equal('five');
      expect(ws2.getCell('D2').value).to.equal(7);
      expect(ws2.getCell('D2').note.texts).to.deep.equal(note.texts);
      expect(ws2.getCell('D2').note.margins).to.deep.equal(note.margins);
      expect(ws2.getCell('D2').note.protection).to.deep.equal(note.protection);
      expect(ws2.getCell('D2').note.editAs).to.deep.equal(note.editAs);
    });

    it('Cell annotation supports setting margins and protection properties', async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('Hello');
      ws.getCell('B2').value = 5;
      ws.getCell('B2').note = 'five';
      const note = {
        texts: [
          {
            font: {
              size: 12,
              color: { argb: 'FFFF6600' },
              name: 'Calibri',
              scheme: 'minor',
            },
            text: 'seven',
          },
        ],
        margins: {
          insetmode: 'custom',
          inset: [0.25, 0.25, 0.35, 0.35],
        },
        protection: {
          locked: 'False',
          lockText: 'False',
        },
        editAs: 'oneCells',
      };
      ws.getCell('D2').value = 7;
      ws.getCell('D2').note = note;

      await wb.commit();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet('Hello');
      expect(ws2.getCell('B2').value).to.equal(5);
      expect(ws2.getCell('B2').note).to.equal('five');

      expect(ws2.getCell('D2').value).to.equal(7);
      expect(ws2.getCell('D2').note.texts).to.deep.equal(note.texts);
      expect(ws2.getCell('D2').note.margins).to.deep.equal(note.margins);
      expect(ws2.getCell('D2').note.protection).to.deep.equal(note.protection);
      expect(ws2.getCell('D2').note.editAs).to.deep.equal(note.editAs);
    });

    it('with background image', async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('Hello');

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });
      ws.getCell('A1').value = 'Hello, World!';
      ws.addBackgroundImage(imageId);

      await wb.commit();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet('Hello');

      const backgroundId2 = ws2.getBackgroundImageId();
      const image = wb2.getImage(backgroundId2);
      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
    });

    it('with background image where worksheet is commited in advance', async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('Hello');

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: 'jpeg',
      });
      ws.getCell('A1').value = 'Hello, World!';
      ws.addBackgroundImage(imageId);

      await ws.commit();
      await wb.commit();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet('Hello');

      const backgroundId2 = ws2.getBackgroundImageId();
      const image = wb2.getImage(backgroundId2);
      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
    });

    describe('image sources', () => {
      const readBackground = async (addImage) => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: IMAGE_SOURCE_FILE });
        const ws = wb.addWorksheet('Hello');
        const imageId = addImage(wb);
        ws.getCell('A1').value = 'Hello, World!';
        ws.addBackgroundImage(imageId);
        await wb.commit();

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.readFile(IMAGE_SOURCE_FILE);
        const ws2 = wb2.getWorksheet('Hello');
        return wb2.getImage(ws2.getBackgroundImageId());
      };

      it('with background image from a buffer', async () => {
        const imageData = await fsReadFileAsync(IMAGE_FILENAME);
        const image = await readBackground((wb) => wb.addImage({ buffer: imageData, extension: 'png' }));
        expect(image.extension).to.equal('png');
        expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
      });

      const base64Image = async () => {
        const imageData = await fsReadFileAsync(IMAGE_FILENAME);
        const base64 = `data:image/png;base64,${imageData.toString('base64')}`;
        const image = await readBackground((wb) => wb.addImage({ base64, extension: 'png' }));
        return { imageData, image };
      };

      it('with background image from base64 keeps the image part and extension', async () => {
        const { image } = await base64Image();
        expect(image.extension).to.equal('png');
        expect(image.buffer).to.be.an.instanceOf(Buffer);
      });

      // archiver ignores the `base64` entry option, so the base64 text is stored instead of the image bytes
      knownBug(
        'with background image from base64 stores the decoded bytes',
        async () => {
          const { imageData, image } = await base64Image();
          expect(image.buffer.length).to.equal(imageData.length);
          expect(Buffer.compare(imageData, image.buffer)).to.equal(0);
        },
        /expected 2232 to equal 1672/,
      );

      it('rejects commit when the image has no filename, buffer or base64', async () => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
          stream: new Writable({
            write(chunk, encoding, callback) {
              callback();
            },
          }),
        });
        const ws = wb.addWorksheet('Hello');
        ws.addBackgroundImage(wb.addImage({ extension: 'png' }));

        let error;
        try {
          await wb.commit();
        } catch (e) {
          error = e;
        }
        expect(error).to.be.an('error');
        expect(error.message).to.equal('Unsupported media');
      });
    });

    it('with conditional formatting', async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true,
        useSharedStrings: true,
      };
      const wb = testUtils.createTestBook(new ExcelJS.stream.xlsx.WorkbookWriter(options), 'xlsx', [
        'conditionalFormatting',
      ]);

      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then((wb2) => {
          testUtils.checkTestBook(wb2, 'xlsx', ['conditionalFormatting']);
        });
    });

    it('with conditional formatting that contains numFmt (#1814)', async () => {
      const sheet = 'conditionalFormatting';
      const options = { filename: TEST_XLSX_FILE_NAME, useStyles: true };

      // generate file with conditional formatting that contains styles with numFmt
      const wb1 = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws1 = wb1.addWorksheet(sheet);
      const cf1 = testUtils.conditionalFormatting.abbreviation;
      ws1.addConditionalFormatting(cf1);
      await wb1.commit();

      // read generated file and extract saved conditional formatting rule
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet(sheet);
      const [cf2] = ws2.conditionalFormattings;

      // verify that rules from generated file contain styles with valid numFmt
      cf2.rules.forEach((rule) => {
        expect(rule.style.numFmt).to.exist;
        expect(rule.style.numFmt.id).to.be.a('number');
        expect(rule.style.numFmt.formatCode).to.be.a('string');
      });
    });
  });
});
