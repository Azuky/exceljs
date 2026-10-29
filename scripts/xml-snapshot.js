/* eslint-disable no-console, no-process-exit, no-await-in-loop, import/no-dynamic-require */
// Pin the timezone first: Dates are rendered in local time (e.g. in CSV), which would make snapshots machine-dependent.
process.env.TZ = 'UTC';
// Writes a set of deterministic workbooks and unpacks their XLSX parts to disk, so output can be
// diffed across toolchains.
// Usage: node scripts/xml-snapshot.js --entry <path-to-exceljs-entry> --out <dir>
const fs = require('fs');
const path = require('path');
const {PassThrough} = require('stream');
const JSZip = require('jszip');

function arg(name) {
  const i = process.argv.indexOf(name);
  return i === -1 ? undefined : process.argv[i + 1];
}
const entry = arg('--entry');
const out = arg('--out');
if (!entry || !out) {
  console.error('Usage: node scripts/xml-snapshot.js --entry <path> --out <dir>');
  process.exit(2);
}
const ExcelJS = require(path.resolve(entry));

const FIXED_DATE = new Date(Date.UTC(2020, 0, 1));
const IMAGE_PATH = path.resolve(__dirname, '../spec/integration/data/image.png');

function newWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.created = FIXED_DATE;
  wb.modified = FIXED_DATE;
  return wb;
}

function stylesAndNumfmt() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Styles');
  ws.getCell('A1').value = 'Bold red';
  ws.getCell('A1').font = {name: 'Arial', size: 14, bold: true, italic: true, color: {argb: 'FFFF0000'}};
  ws.getCell('A2').value = 'Solid fill';
  ws.getCell('A2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
  ws.getCell('A3').value = 'Gradient';
  ws.getCell('A3').fill = {
    type: 'gradient',
    gradient: 'angle',
    degree: 0,
    stops: [
      {position: 0, color: {argb: 'FF0000FF'}},
      {position: 1, color: {argb: 'FFFFFF00'}},
    ],
  };
  ws.getCell('A4').value = 'Borders';
  ws.getCell('A4').border = {
    top: {style: 'thin', color: {argb: 'FF000000'}},
    left: {style: 'double'},
    bottom: {style: 'thick', color: {argb: 'FF0000FF'}},
    right: {style: 'dashed'},
  };
  ws.getCell('A5').value = 'Aligned wrapped text that is long';
  ws.getCell('A5').alignment = {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
    indent: 1,
    textRotation: 45,
  };
  ws.getCell('B1').value = 1234.5678;
  ws.getCell('B1').numFmt = '#,##0.00';
  ws.getCell('B2').value = 0.256;
  ws.getCell('B2').numFmt = '0.00%';
  ws.getCell('B3').value = new Date(Date.UTC(2021, 5, 15, 12, 30));
  ws.getCell('B3').numFmt = 'yyyy-mm-dd hh:mm';
  ws.getCell('B4').value = new Date(Date.UTC(2021, 11, 31));
  ws.getCell('B5').value = 42;
  ws.getCell('B5').numFmt = '"$"#,##0;[Red]-"$"#,##0';
  ws.getCell('B5').protection = {locked: false, hidden: true};
  return wb;
}

function formulas() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Formulas');
  ws.getCell('A1').value = 1;
  ws.getCell('A2').value = 2;
  ws.getCell('A3').value = 3;
  ws.getCell('B1').value = {formula: 'SUM(A1:A3)'};
  ws.getCell('B2').value = {formula: 'SUM(A1:A3)*2', result: 12};
  ws.getCell('B3').value = {formula: 'A1&"x"', result: '1x'};
  ws.getCell('C1').value = {formula: 'A1*10', result: 10};
  ws.getCell('C2').value = {sharedFormula: 'C1', result: 20};
  ws.getCell('C3').value = {sharedFormula: 'C1', result: 30};
  ws.getCell('D1').value = {formula: 'A1:A3*2', result: 2, shareType: 'array', ref: 'D1:D3'};
  ws.getCell('D2').value = 4;
  ws.getCell('D3').value = 6;
  ws.getCell('E1').value = {formula: 'A1>0', result: true};
  ws.getCell('E2').value = {formula: 'NA()', result: {error: '#N/A'}};
  return wb;
}

function mergesAndSizes() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Merges');
  ws.mergeCells('A1:C2');
  ws.getCell('A1').value = 'Merged';
  ws.mergeCells('E5', 'F7');
  ws.getColumn(1).width = 30;
  ws.getColumn(2).width = 5.5;
  ws.getColumn(3).hidden = true;
  ws.getRow(1).height = 42.5;
  ws.getRow(4).hidden = true;
  ws.getCell('A4').value = 'hidden row';
  ws.getRow(6).height = 10;
  ws.getCell('A6').value = 'short row';
  return wb;
}

function table() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Table');
  ws.addTable({
    name: 'MyTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: {theme: 'TableStyleDark3', showRowStripes: true},
    columns: [
      {name: 'Date', totalsRowLabel: 'Totals:', filterButton: true, style: {numFmt: 'yyyy-mm-dd'}},
      {
        name: 'Amount',
        totalsRowFunction: 'sum',
        filterButton: false,
        style: {font: {bold: true}, numFmt: '0.00'},
      },
    ],
    rows: [
      [new Date(Date.UTC(2019, 6, 20)), 70.1],
      [new Date(Date.UTC(2019, 6, 21)), 70.6],
      [new Date(Date.UTC(2019, 6, 22)), 70.1],
    ],
  });
  return wb;
}

function image() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Image');
  const id = wb.addImage({filename: IMAGE_PATH, extension: 'png'});
  ws.addImage(id, 'B2:D6');
  ws.addImage(id, {tl: {col: 5, row: 1}, ext: {width: 100, height: 80}});
  return wb;
}

function validationAndCf() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Validation');
  ws.getCell('A1').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"One,Two,Three,Four"'],
    showErrorMessage: true,
    errorTitle: 'Bad',
    error: 'Pick from the list',
  };
  ws.getCell('A2').dataValidation = {
    type: 'whole',
    operator: 'between',
    allowBlank: true,
    formulae: [1, 10],
    showInputMessage: true,
    promptTitle: 'Number',
    prompt: 'A whole number 1-10',
  };
  for (let r = 1; r <= 7; r++) {
    for (let c = 3; c <= 7; c++) ws.getCell(r, c).value = r * c;
  }
  ws.addConditionalFormatting({
    ref: 'C1:G7',
    rules: [
      {
        type: 'expression',
        formulae: ['MOD(ROW()+COLUMN(),2)=0'],
        style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FF00FF00'}}},
      },
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: [20],
        style: {font: {color: {argb: 'FFFF0000'}}},
      },
    ],
  });
  ws.addConditionalFormatting({
    ref: 'I1:I7',
    rules: [
      {
        type: 'colorScale',
        cfvo: [{type: 'min'}, {type: 'max'}],
        color: [{argb: 'FFFFFFFF'}, {argb: 'FF00AA00'}],
      },
    ],
  });
  return wb;
}

function commentsNamesLinks() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Misc');
  ws.getCell('A1').value = 'With note';
  ws.getCell('A1').note = 'Hello, ExcelJS!';
  ws.getCell('A2').value = 'Rich note';
  ws.getCell('A2').note = {
    texts: [
      {font: {size: 12, color: {argb: 'FFFF0000'}, name: 'Calibri'}, text: 'Red '},
      {font: {size: 10, name: 'Calibri'}, text: 'plain'},
    ],
    editAs: 'twoCells',
  };
  ws.getCell('B1').value = 100;
  wb.definedNames.add('Misc!$B$1', 'MyName');
  ws.getCell('C1').value = {text: 'ExcelJS', hyperlink: 'https://github.com/exceljs/exceljs'};
  ws.getCell('C2').value = {text: 'Internal', hyperlink: '#Misc!A1'};
  return wb;
}

async function streamWriter() {
  const stream = new PassThrough();
  const chunks = [];
  stream.on('data', chunk => chunks.push(chunk));
  const done = new Promise((resolve, reject) => {
    stream.on('end', resolve);
    stream.on('error', reject);
  });
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream,
    useStyles: true,
    useSharedStrings: true,
  });
  wb.created = FIXED_DATE;
  wb.modified = FIXED_DATE;
  const ws = wb.addWorksheet('Stream');
  ws.columns = [
    {header: 'Name', key: 'name', width: 20},
    {header: 'Amount', key: 'amount', width: 12, style: {numFmt: '#,##0.00'}},
    {header: 'When', key: 'when', width: 14, style: {numFmt: 'yyyy-mm-dd'}},
  ];
  ws.mergeCells('E1:F2');
  ws.getRow(1).font = {bold: true, color: {argb: 'FFFFFFFF'}};
  ws.getRow(1).fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF333333'}};
  ws.getRow(1).commit();
  for (let i = 1; i <= 5; i++) {
    const row = ws.addRow({name: `Item ${i}`, amount: i * 1.5, when: new Date(Date.UTC(2020, i, 1))});
    row.getCell('name').border = {bottom: {style: 'thin'}};
    row.commit();
  }
  ws.commit();
  await wb.commit();
  await done;
  return Buffer.concat(chunks);
}

async function unpack(buffer, dir) {
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files)
    .filter(n => !zip.files[n].dir)
    .sort();
  for (const name of names) {
    const target = path.join(dir, name);
    fs.mkdirSync(path.dirname(target), {recursive: true});
    fs.writeFileSync(target, await zip.files[name].async('nodebuffer'));
  }
}

const workbookScenarios = {
  'styles-and-numfmt': stylesAndNumfmt,
  formulas,
  'merges-and-sizes': mergesAndSizes,
  table,
  image,
  'validation-and-cf': validationAndCf,
  'comments-names-links': commentsNamesLinks,
};

async function main() {
  if (fs.existsSync(out) && fs.readdirSync(out).length > 0) {
    console.error(`--out directory is not empty: ${out}. Remove it or pass a fresh directory.`);
    process.exit(2);
  }
  fs.mkdirSync(out, {recursive: true});
  for (const [name, build] of Object.entries(workbookScenarios)) {
    const buffer = await build().xlsx.writeBuffer();
    await unpack(buffer, path.join(out, name));
  }
  await unpack(await streamWriter(), path.join(out, 'stream-writer'));

  const csvWb = newWorkbook();
  const csvWs = csvWb.addWorksheet('Csv');
  csvWs.addRows([
    ['a', 'b, with comma', 'c "quoted"'],
    [1, 2.5, new Date(Date.UTC(2020, 0, 1))],
    [true, null, {formula: 'A2+B2', result: 3.5}],
  ]);
  const csvDir = path.join(out, 'csv-write');
  fs.mkdirSync(csvDir, {recursive: true});
  fs.writeFileSync(path.join(csvDir, 'out.csv'), await csvWb.csv.writeBuffer());
  console.log(`snapshot written to ${out}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
