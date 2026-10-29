const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 163 - Error while using xslx readFile method', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx.readFile('./spec/integration/data/test-issue-163.xlsx').then(() => {
      const ws = wb.getWorksheet('Sheet1');
      expect([1, 2, 3].map((r) => ws.getCell(`A${r}`).value)).to.deep.equal([1, 2, 3]);
    });
  });
});
