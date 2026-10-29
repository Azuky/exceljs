const CellMatrix = verquire('utils/cell-matrix');

describe('CellMatrix', () => {
  it('getCell always returns a cell', () => {
    const cm = new CellMatrix();
    expect(cm.getCell('A1')).to.include({ address: 'A1', row: 1, col: 1 });
    expect(cm.getCell('$B$2')).to.include({ address: 'B2', row: 2, col: 2 });
    expect(cm.getCell('Sheet!$C$3')).to.include({ address: 'C3', row: 3, col: 3, sheetName: 'Sheet' });
  });
  it('getCell returns the same instance as findCell afterwards', () => {
    const cm = new CellMatrix();
    const cell = cm.getCell('$B$2');
    expect(cm.getCell('B2')).to.equal(cell);
    expect(cm.findCell('B2')).to.equal(cell);
  });
  it('keeps cells of different sheets apart', () => {
    const cm = new CellMatrix();
    const plain = cm.getCell('A1');
    const other = cm.getCell('Sheet!A1');
    expect(other).not.to.equal(plain);
    expect(cm.findCell('Sheet!A1')).to.equal(other);
    expect(cm.findCell('Other!A1')).to.be.undefined;
    expect(cm.findCellAt('Sheet', 1, 1)).to.equal(other);
  });
  it('findCell only returns known cells', () => {
    const cm = new CellMatrix();
    expect(cm.findCell('A1')).to.be.undefined;
    expect(cm.findCell('$B$2')).to.be.undefined;
    expect(cm.findCell('Sheet!$C$3')).to.be.undefined;
  });
});
