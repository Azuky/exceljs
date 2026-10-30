const fs = require('fs');
const JSZip = require('jszip');
const tmp = require('tmp');

const { knownBug } = require('../utils/known-bug');

const ExcelJS = verquire('exceljs');
const HyperlinkReader = verquire('stream/xlsx/hyperlink-reader');
const RelType = verquire('xlsx/rel-type');

const { WorkbookReader } = ExcelJS.stream.xlsx;

const OUT = './spec/out';
const file = (name) => `${OUT}/wb-stream-reader.${name}.xlsx`;

// resolves on 'end', rejects on 'error' (and fails if both would be needed)
const run = (reader) =>
  new Promise((resolve, reject) => {
    reader.on('end', resolve);
    reader.on('error', reject);
    reader.read();
  });

// reads every row of every worksheet through the event interface
const readSheets = async (filename, options) => {
  const reader = new WorkbookReader(filename, options);
  const sheets = [];
  reader.on('worksheet', (worksheet) => {
    const sheet = { name: worksheet.name, rows: [] };
    sheets.push(sheet);
    worksheet.on('row', (row) => sheet.rows.push(row));
  });
  await run(reader);
  return sheets;
};

const writeWorkbook = async (name, fill) => {
  const wb = new ExcelJS.Workbook();
  fill(wb.addWorksheet('Sheet'), wb);
  const filename = file(name);
  await wb.xlsx.writeFile(filename);
  return filename;
};

const FULL_FONT = {
  name: 'Arial',
  size: 14,
  family: 2,
  charset: 1,
  underline: true,
  vertAlign: 'superscript',
  color: { argb: 'FFFF0000' },
  bold: true,
  italic: true,
  strike: true,
  outline: true,
};

describe('WorkbookReader (streaming)', () => {
  describe('shared strings', () => {
    let filename;
    beforeAll(async () => {
      filename = await writeWorkbook('shared-strings', (ws) => {
        ws.getCell('A1').value = { richText: [{ text: 'styled ', font: FULL_FONT }, { text: 'plain' }] };
        ws.getCell('B1').value = 'plain';
        ws.getCell('C1').value = 'other';
      });
    });

    it('sharedStrings:emit leaves cell values as references to the string index', async () => {
      const [sheet] = await readSheets(filename, { sharedStrings: 'emit' });
      expect(sheet.rows.map((row) => row.values)).to.deep.equal([
        [undefined, { sharedString: 0 }, { sharedString: 1 }, { sharedString: 2 }],
      ]);
    });

    it('sharedStrings:ignore leaves cell values as references to the string index', async () => {
      const [sheet] = await readSheets(filename, { sharedStrings: 'ignore' });
      expect(sheet.rows.map((row) => row.values)).to.deep.equal([
        [undefined, { sharedString: 0 }, { sharedString: 1 }, { sharedString: 2 }],
      ]);
    });

    it('sharedStrings:cache inserts the strings into the cells', async () => {
      const [sheet] = await readSheets(filename, { sharedStrings: 'cache' });
      const [, rich, plain, other] = sheet.rows[0].values;
      expect(rich.richText.map((run) => run.text)).to.deep.equal(['styled ', 'plain']);
      expect(plain).to.equal('plain');
      expect(other).to.equal('other');
    });

    // README: the workbook emits 'shared-strings'. _parseSharedStrings yields bare { index, text } objects,
    // but read() destructures { eventType, value }, so nothing is ever emitted.
    knownBug(
      "sharedStrings:emit emits a 'shared-strings' event per string",
      async () => {
        const reader = new WorkbookReader(filename, { sharedStrings: 'emit', worksheets: 'emit' });
        const events = [];
        reader.on('shared-strings', (value) => events.push(value));
        await run(reader);

        expect(events.map((event) => event.index)).to.deep.equal([0, 1, 2]);
        const texts = events.map((event) =>
          typeof event.text === 'string' ? event.text : event.text.richText.map((r) => r.text).join(''),
        );
        expect(texts).to.deep.equal(['styled plain', 'plain', 'other']);
      },
      /expected \[\] to deeply equal \[ \+?0, 1, 2 \]/,
    );

    knownBug(
      "parse() yields 'shared-strings' events",
      async () => {
        const reader = new WorkbookReader(filename, { sharedStrings: 'emit', worksheets: 'emit' });
        const eventTypes = [];
        for await (const { eventType } of reader.parse()) {
          eventTypes.push(eventType);
        }
        expect(eventTypes.filter((type) => type === 'shared-strings')).to.have.length(3);
      },
      /to have a length of 3 but got \+?0/,
    );
  });

  describe('rich text fonts', () => {
    let font;
    let themeFont;
    beforeAll(async () => {
      const filename = await writeWorkbook('rich-text', (ws) => {
        ws.getCell('A1').value = { richText: [{ text: 'styled ', font: FULL_FONT }, { text: 'plain' }] };
        ws.getCell('A2').value = { richText: [{ text: 'themed', font: { color: { theme: 1 } } }] };
      });
      const [sheet] = await readSheets(filename, { sharedStrings: 'cache' });
      ({ font } = sheet.rows[0].values[1].richText[0]);
      ({ font: themeFont } = sheet.rows[1].values[1].richText[0]);
    });

    it('reads bold, italic, underline, outline, family, size and vertAlign', () => {
      expect(font).to.deep.include({
        bold: true,
        italic: true,
        underline: true,
        outline: true,
        family: 2,
        size: 14,
        vertAlign: 'superscript',
      });
    });

    it('reads the text of every run, formatted or not', async () => {
      const filename = await writeWorkbook('rich-text-plain-run', (ws) => {
        ws.getCell('A1').value = { richText: [{ text: 'styled ', font: FULL_FONT }, { text: 'plain' }] };
      });
      const [sheet] = await readSheets(filename, { sharedStrings: 'cache' });
      const { richText } = sheet.rows[0].values[1];
      expect(richText.map((run) => run.text)).to.deep.equal(['styled ', 'plain']);
    });

    knownBug(
      'reads the rgb colour of a run as argb',
      () => {
        expect(font.color.argb).to.equal('FFFF0000');
      },
      /expected undefined to equal 'FFFF0000'/,
    );

    knownBug(
      'reads the theme colour of a run as a number',
      () => {
        expect(themeFont.color.theme).to.equal(1);
      },
      /expected '1' to equal 1/,
    );

    // _parseSharedStrings reads node.value instead of node.attributes.val
    knownBug(
      'reads the font name of a run',
      () => {
        expect(font.name).to.equal('Arial');
      },
      /expected undefined to equal 'Arial'/,
    );

    // _parseSharedStrings reads node.attributes.charset instead of node.attributes.val
    knownBug(
      'reads the charset of a run',
      () => {
        expect(font.charset).to.equal(1);
      },
      /expected NaN to equal 1/,
    );

    // the 'strike' case is an empty break
    knownBug(
      'reads strike of a run',
      () => {
        expect(font.strike).to.equal(true);
      },
      /expected undefined to equal true/,
    );
  });

  describe('hyperlinks', () => {
    let filename;
    beforeAll(async () => {
      filename = await writeWorkbook('hyperlinks', (ws) => {
        ws.getCell('A1').value = 'before';
        ws.getCell('B1').value = { text: 'link', hyperlink: 'http://x.com' };
      });
    });

    it("hyperlinks:emit emits a hyperlinks reader whose read() emits 'hyperlink'", async () => {
      const reader = new WorkbookReader(filename, { hyperlinks: 'emit' });
      const hyperlinks = [];
      const sheetLinks = [];
      const finished = [];
      reader.on('hyperlinks', (hyperlinksReader) => {
        hyperlinksReader.on('hyperlink', (hyperlink) => hyperlinks.push(hyperlink));
        finished.push(new Promise((resolve) => hyperlinksReader.on('finished', resolve)));
        hyperlinksReader.read();
      });
      const rows = [];
      reader.on('worksheet', (worksheet) => {
        worksheet.on('hyperlink', (hyperlink) => sheetLinks.push(hyperlink));
        worksheet.on('row', (row) => rows.push(row.values));
      });
      await run(reader);
      await Promise.all(finished);

      expect(hyperlinks).to.have.length(1);
      expect(hyperlinks[0]).to.include({ rId: 'rId1', target: 'http://x.com', targetMode: 'External' });
      expect(sheetLinks).to.deep.equal([{ ref: 'B1', rId: 'rId1' }]);
      expect(rows).to.deep.equal([[undefined, 'before', 'link']]);
    });

    it('hyperlinks:cache keeps the worksheet hyperlinks by cell reference', async () => {
      const reader = new WorkbookReader(filename, { hyperlinks: 'cache' });
      let worksheetReader;
      reader.on('worksheet', (worksheet) => {
        worksheetReader = worksheet;
      });
      await run(reader);

      expect(worksheetReader.hyperlinks).to.deep.equal({ B1: { ref: 'B1', rId: 'rId1' } });
    });

    it('hyperlinks:ignore emits nothing and keeps nothing', async () => {
      const reader = new WorkbookReader(filename, { hyperlinks: 'ignore' });
      let emitted = 0;
      let worksheetReader;
      reader.on('hyperlinks', () => emitted++);
      reader.on('worksheet', (worksheet) => {
        worksheet.on('hyperlink', () => emitted++);
        worksheetReader = worksheet;
      });
      await run(reader);

      expect(emitted).to.equal(0);
      expect(worksheetReader.hyperlinks).to.equal(undefined);
    });

    describe('HyperlinkReader', () => {
      const RELS = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        `<Relationship Id="rId1" Type="${RelType.Hyperlink}" Target="http://x.com" TargetMode="External"/>`,
        `<Relationship Id="rId2" Type="${RelType.Hyperlink}" Target="http://y.com" TargetMode="External"/>`,
        `<Relationship Id="rId3" Type="${RelType.VmlDrawing}" Target="../drawings/vmlDrawing1.vml"/>`,
        '</Relationships>',
      ].join('');

      const createReader = (options) =>
        new HyperlinkReader({
          workbook: null,
          id: '1',
          iterator: (async function* iterator() {
            yield Buffer.from(RELS);
          })(),
          options,
        });

      it('caches hyperlink relationships by rId and skips other relationships', async () => {
        const reader = createReader({ hyperlinks: 'cache' });
        await reader.read();
        expect(Object.keys(reader.hyperlinks)).to.deep.equal(['rId1', 'rId2']);
        expect(reader.hyperlinks.rId1).to.include({ rId: 'rId1', target: 'http://x.com', targetMode: 'External' });
        expect(reader.hyperlinks.rId2).to.include({ rId: 'rId2', target: 'http://y.com', targetMode: 'External' });
      });

      it('emits one hyperlink per hyperlink relationship', async () => {
        const reader = createReader({ hyperlinks: 'emit' });
        const targets = [];
        reader.on('hyperlink', (hyperlink) => targets.push(hyperlink.target));
        await reader.read();
        expect(targets).to.deep.equal(['http://x.com', 'http://y.com']);
      });

      it('does nothing but finish when hyperlinks are ignored', async () => {
        const reader = createReader({ hyperlinks: 'ignore' });
        let finished = false;
        reader.on('finished', () => {
          finished = true;
        });
        await reader.read();
        expect(finished).to.equal(true);
        expect(reader.count).to.equal(0);
      });

      // the cache is an object keyed by rId, but count reads .length and each calls .forEach
      knownBug(
        'count is the number of cached hyperlinks',
        async () => {
          const reader = createReader({ hyperlinks: 'cache' });
          await reader.read();
          expect(reader.count).to.equal(2);
        },
        /expected \+?0 to equal 2/,
      );

      knownBug(
        'each iterates the cached hyperlinks',
        async () => {
          const reader = createReader({ hyperlinks: 'cache' });
          await reader.read();
          const targets = [];
          reader.each((hyperlink) => targets.push(hyperlink.target));
          expect(targets).to.deep.equal(['http://x.com', 'http://y.com']);
        },
        /forEach is not a function/,
      );
    });
  });

  describe('string cells written by other producers', () => {
    const rewriteSheet = async (name, replacements) => {
      const filename = file(name);
      const source = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file(`${name}-source`) });
      const ws = source.addWorksheet('Sheet');
      ws.addRow(['str cell', { formula: '"a"&"b"', result: 'ab' }, 'INLINE']).commit();
      ws.commit();
      await source.commit();

      const zip = await JSZip.loadAsync(fs.readFileSync(file(`${name}-source`)));
      let xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
      for (const [from, to] of replacements) {
        expect(xml).to.include(from);
        xml = xml.replace(from, to);
      }
      zip.file('xl/worksheets/sheet1.xml', xml);
      fs.writeFileSync(filename, await zip.generateAsync({ type: 'nodebuffer' }));
      return filename;
    };

    it('reads t="str" cells, string formula results and inline strings', async () => {
      const filename = await rewriteSheet('inline', [
        ['<c r="C1" t="str"><v>INLINE</v></c>', '<c r="C1" t="inlineStr"><is><t>inline text</t></is></c>'],
      ]);
      const [sheet] = await readSheets(filename, { sharedStrings: 'ignore' });
      expect(sheet.rows.map((row) => row.values)).to.deep.equal([
        [undefined, 'str cell', { formula: '"a"&"b"', result: 'ab' }, 'inline text'],
      ]);
    });

    it('decodes XML special characters in t="str" cells', async () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file('str-plain') });
      const ws = wb.addWorksheet('Sheet');
      ws.addRow(['x & y <z>']).commit();
      ws.commit();
      await wb.commit();

      const [sheet] = await readSheets(file('str-plain'), { sharedStrings: 'ignore' });
      expect(sheet.rows[0].values).to.deep.equal([undefined, 'x & y <z>']);
    });

    // the text is already decoded by the XML parser, then utils.xmlDecode decodes it again
    knownBug(
      'does not decode entities twice',
      async () => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file('str-entity') });
        const ws = wb.addWorksheet('Sheet');
        ws.addRow(['Tom &amp; Jerry']).commit();
        ws.commit();
        await wb.commit();

        const [sheet] = await readSheets(file('str-entity'), { sharedStrings: 'ignore' });
        expect(sheet.rows[0].values[1]).to.equal('Tom &amp; Jerry');
      },
      /expected 'Tom & Jerry' to equal 'Tom &amp; Jerry'/,
    );
  });

  describe('string formula results', () => {
    // same double decoding as above, on the formula result branch
    knownBug(
      'does not decode entities twice in a string formula result',
      async () => {
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file('formula-entity') });
        const ws = wb.addWorksheet('Sheet');
        ws.addRow([{ formula: '"R &amp; D"', result: 'R &amp; D' }]).commit();
        ws.commit();
        await wb.commit();

        const [sheet] = await readSheets(file('formula-entity'), { sharedStrings: 'ignore' });
        expect(sheet.rows[0].values[1].result).to.equal('R &amp; D');
      },
      /expected 'R & D' to equal 'R &amp; D'/,
    );
  });

  describe('errors', () => {
    const readFailure = async (reader) => {
      let ended = false;
      reader.on('end', () => {
        ended = true;
      });
      let error;
      try {
        await run(reader);
      } catch (e) {
        error = e;
      }
      expect(ended).to.equal(false);
      return error;
    };

    it('emits error for an input that is neither a path nor a stream', async () => {
      const error = await readFailure(new WorkbookReader({}));
      expect(error).to.be.an('error');
      expect(error.message).to.equal('Could not recognise input: [object Object]');
    });

    it('emits error for a file that is not a zip archive', async () => {
      fs.writeFileSync(file('not-a-zip'), 'this is not a zip archive');
      const error = await readFailure(new WorkbookReader(file('not-a-zip')));
      expect(error).to.be.an('error');
      expect(error.message).to.match(/invalid signature/);
    });

    it('emits error for a truncated zip archive', async () => {
      const filename = await writeWorkbook('to-truncate', (ws) => {
        ws.getCell('A1').value = 'hello';
      });
      const bytes = fs.readFileSync(filename);
      fs.writeFileSync(file('truncated'), bytes.subarray(0, bytes.length >> 1));
      const error = await readFailure(new WorkbookReader(file('truncated')));
      expect(error).to.be.an('error');
      expect(error.message).to.match(/FILE_ENDED/);
    });

    // Without the stream error listener installed here the 'error' of the missing file is an uncaught
    // exception on the fs stream; with it, the reader never reports anything.
    knownBug(
      "emits 'error' for a file that does not exist",
      async () => {
        const createReadStream = fs.createReadStream.bind(fs);
        const spy = vi.spyOn(fs, 'createReadStream').mockImplementation((...args) => {
          const stream = createReadStream(...args);
          stream.on('error', () => {});
          return stream;
        });
        try {
          const reader = new WorkbookReader(file('does-not-exist'));
          const error = await new Promise((resolve, reject) => {
            const timer = setTimeout(() => reject(new Error("no 'error' event for a missing file")), 500);
            reader.on('error', (e) => {
              clearTimeout(timer);
              resolve(e);
            });
            reader.on('end', () => {
              clearTimeout(timer);
              reject(new Error("'end' emitted for a missing file"));
            });
            reader.read();
          });
          expect(error.code).to.equal('ENOENT');
        } finally {
          spy.mockRestore();
        }
      },
      /no 'error' event for a missing file/,
    );
  });

  describe('options and zip entry order', () => {
    const CONTAINER_FIRST = [
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/_rels/workbook.xml.rels',
      'xl/workbook.xml',
      'xl/sharedStrings.xml',
      'xl/styles.xml',
      'xl/worksheets/sheet1.xml',
    ];

    // rebuilds the zip with everything the worksheet needs placed before it
    const reorder = async (source, target) => {
      const zip = await JSZip.loadAsync(fs.readFileSync(source));
      const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
      const ordered = [...CONTAINER_FIRST, ...names.filter((name) => !CONTAINER_FIRST.includes(name))];
      const out = new JSZip();
      for (const name of ordered) {
        out.file(name, await zip.files[name].async('nodebuffer'));
      }
      fs.writeFileSync(target, await out.generateAsync({ type: 'nodebuffer' }));
    };

    let filename;
    let ordered;
    beforeAll(async () => {
      filename = await writeWorkbook('options', (ws) => {
        ws.name = 'Named Sheet';
        ws.addRow(['alpha', 'beta', 42]);
        ws.addRow(['gamma', 'alpha', 7]);
      });
      ordered = file('options-ordered');
      await reorder(filename, ordered);
    });

    const EXPECTED_ROWS = [
      [undefined, 'alpha', 'beta', 42],
      [undefined, 'gamma', 'alpha', 7],
    ];

    it('streams a worksheet straight through when everything it needs precedes it', async () => {
      const tmpFile = vi.spyOn(tmp, 'file');
      try {
        const sheets = await readSheets(ordered, { sharedStrings: 'cache' });
        expect(sheets).to.have.length(1);
        expect(sheets[0].name).to.equal('Named Sheet');
        expect(sheets[0].rows.map((row) => row.values)).to.deep.equal(EXPECTED_ROWS);
        expect(tmpFile).not.toHaveBeenCalled();
      } finally {
        tmpFile.mockRestore();
      }
    });

    it('defers the worksheet through a temp file when shared strings are ignored', async () => {
      const tmpFile = vi.spyOn(tmp, 'file');
      try {
        const sheets = await readSheets(ordered, { sharedStrings: 'ignore' });
        expect(sheets).to.have.length(1);
        expect(sheets[0].name).to.equal('Named Sheet');
        expect(sheets[0].rows.map((row) => row.values)).to.deep.equal([
          [undefined, { sharedString: 0 }, { sharedString: 1 }, 42],
          [undefined, { sharedString: 2 }, { sharedString: 0 }, 7],
        ]);
        expect(tmpFile).toHaveBeenCalledTimes(1);
      } finally {
        tmpFile.mockRestore();
      }
    });

    it('reads the same rows whether or not the worksheet had to be deferred', async () => {
      const [deferred] = await readSheets(filename, { sharedStrings: 'cache' });
      expect(deferred.name).to.equal('Named Sheet');
      expect(deferred.rows.map((row) => row.values)).to.deep.equal(EXPECTED_ROWS);
    });

    it('worksheets:ignore emits no worksheet but still ends', async () => {
      const reader = new WorkbookReader(filename, { worksheets: 'ignore' });
      let worksheets = 0;
      reader.on('worksheet', () => worksheets++);
      await run(reader);
      expect(worksheets).to.equal(0);
    });

    // delays the real fs.open so a premature temp file cleanup would win the race against the read stream
    const withDelayedOpen = async (fn) => {
      const realOpen = fs.open;
      const open = vi.spyOn(fs, 'open').mockImplementation((...args) => {
        setTimeout(() => realOpen.apply(fs, args), 20);
      });
      try {
        await fn();
      } finally {
        open.mockRestore();
      }
    };

    it('worksheets:ignore survives a slow temp file open without an uncaught error', async () => {
      await withDelayedOpen(async () => {
        const reader = new WorkbookReader(filename, { worksheets: 'ignore' });
        let worksheets = 0;
        reader.on('worksheet', () => worksheets++);
        await run(reader);
        expect(worksheets).to.equal(0);
        // a late ENOENT from a stream opened after cleanup would surface here
        await new Promise((resolve) => setTimeout(resolve, 100));
      });
    });

    it('worksheets:emit keeps the temp file until read, then removes it', async () => {
      const paths = [];
      const realFile = tmp.file;
      const tmpFile = vi.spyOn(tmp, 'file').mockImplementation((...args) => {
        const callback = args[args.length - 1];
        return realFile.apply(tmp, [
          ...args.slice(0, -1),
          (err, path, ...rest) => {
            if (path) paths.push(path);
            return callback(err, path, ...rest);
          },
        ]);
      });
      try {
        await withDelayedOpen(async () => {
          const sheets = await readSheets(filename, { sharedStrings: 'cache' });
          expect(sheets).to.have.length(1);
          expect(sheets[0].rows.map((row) => row.values)).to.deep.equal(EXPECTED_ROWS);
          await new Promise((resolve) => setTimeout(resolve, 100));
        });
        expect(paths).to.have.length(1);
        expect(fs.existsSync(paths[0])).to.equal(false);
      } finally {
        tmpFile.mockRestore();
      }
    });

    it('removes the temp file even when the consumer never reads the worksheet', async () => {
      const paths = [];
      const realFile = tmp.file;
      const tmpFile = vi.spyOn(tmp, 'file').mockImplementation((...args) => {
        const callback = args[args.length - 1];
        return realFile.apply(tmp, [
          ...args.slice(0, -1),
          (err, path, ...rest) => {
            if (path) paths.push(path);
            return callback(err, path, ...rest);
          },
        ]);
      });
      try {
        // eslint-disable-next-line no-unused-vars
        for await (const worksheet of new WorkbookReader(filename, {})) {
          // deliberately not reading the worksheet
        }
        await new Promise((resolve) => setTimeout(resolve, 100));
        expect(paths).to.have.length(1);
        expect(fs.existsSync(paths[0])).to.equal(false);
      } finally {
        tmpFile.mockRestore();
      }
    });

    it('entries:emit reports the parts being read', async () => {
      const reader = new WorkbookReader(filename, { entries: 'emit' });
      const entries = [];
      reader.on('entry', (entry) => entries.push(entry));
      await run(reader);

      expect(entries).to.deep.include({ type: 'workbook' });
      expect(entries).to.deep.include({ type: 'shared-strings' });
      expect(entries).to.deep.include({ type: 'styles' });
      expect(entries).to.deep.include({ type: 'worksheet', id: '1' });
    });

    it('entries:ignore emits no entries', async () => {
      const reader = new WorkbookReader(filename, { entries: 'ignore' });
      let entries = 0;
      reader.on('entry', () => entries++);
      await run(reader);
      expect(entries).to.equal(0);
    });
  });

  describe('styles', () => {
    let filename;
    beforeAll(async () => {
      filename = await writeWorkbook('styles', (ws) => {
        const row = ws.getRow(1);
        row.font = { bold: true };
        row.getCell(1).value = 5;
        row.getCell(1).numFmt = '0.00';
        row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
        row.getCell(1).alignment = { horizontal: 'center' };
        row.getCell(2).value = 'text';
        ws.getRow(2).getCell(1).value = 1;
      });
    });

    it('styles:cache maps row and cell styles', async () => {
      const [sheet] = await readSheets(filename, { styles: 'cache' });
      const [styledRow, plainRow] = sheet.rows;

      expect(styledRow.style.font).to.deep.equal({ bold: true });

      const styledCell = styledRow.getCell(1);
      expect(styledCell.numFmt).to.equal('0.00');
      expect(styledCell.fill).to.deep.equal({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } });
      expect(styledCell.alignment).to.deep.equal({ horizontal: 'center' });
      expect(styledCell.font).to.deep.equal({ bold: true });

      expect(plainRow.style).to.deep.equal({});
      expect(plainRow.getCell(1).style).to.deep.equal({});
    });

    it('styles:ignore leaves rows and cells unstyled', async () => {
      const [sheet] = await readSheets(filename, { styles: 'ignore' });
      const [styledRow] = sheet.rows;

      expect(styledRow.style).to.deep.equal({});
      expect(styledRow.getCell(1).style).to.deep.equal({});
      expect(styledRow.getCell(1).value).to.equal(5);
    });
  });
});
