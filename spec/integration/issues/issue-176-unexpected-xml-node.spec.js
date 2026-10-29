const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 176 - Unexpected xml node in parseOpen', () => {
    const wb = new ExcelJS.Workbook();
    return wb.xlsx.readFile('./spec/integration/data/test-issue-176.xlsx').then(() => {
      const ws = wb.getWorksheet('Sheet');
      expect(ws.getRow(1).values).to.deep.equal([undefined, 'Name', 'Rating']);
      expect(ws.getCell('A4').value).to.equal('Third');
      expect(ws.getCell('B4').value).to.equal(2);
    });
  });
});
