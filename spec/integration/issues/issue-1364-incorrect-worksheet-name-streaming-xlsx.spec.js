const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/integration/data/test-issue-1364.xlsx';

describe('github issues', () => {
  it('issue 1364 - Incorrect Worksheet Name on Streaming XLSX Reader', async () => {
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(TEST_XLSX_FILE_NAME, {});
    const names = [];
    await new Promise((resolve, reject) => {
      workbookReader.on('worksheet', (worksheet) => {
        names.push(worksheet.name);
      });
      workbookReader.on('end', resolve);
      workbookReader.on('error', reject);
      workbookReader.read();
    });
    expect(names).to.deep.equal(['Sum Worksheet']);
  });
});
