const ExcelJS = verquire('exceljs');

// this file to contain integration tests created from github issues

describe('github issues', () => {
  it('issue 880 - malformed comment crashes on write', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('./spec/integration/data/test-issue-880.xlsx');
    const buffer = await wb.xlsx.writeBuffer({
      useStyles: true,
      useSharedStrings: true,
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    expect(wb2.worksheets.length).to.equal(wb.worksheets.length);
    let notesSeen = 0;
    wb.eachSheet((ws) => {
      const ws2 = wb2.getWorksheet(ws.name);
      expect(ws2, ws.name).to.be.ok;
      ws.eachRow((row) => {
        row.eachCell((cell) => {
          if (cell.note) {
            notesSeen++;
            const joined = (note) => note.texts.map((t) => t.text).join('');
            expect(joined(ws2.getCell(cell.address).note), cell.address).to.equal(joined(cell.note));
          }
        });
      });
    });
    expect(notesSeen).to.be.greaterThan(0);
  }, 6000);
});
