const DefinedNames = verquire('doc/defined-names');

describe('DefinedNames', () => {
  it('adds names for cells', () => {
    const dn = new DefinedNames();

    dn.add('blort!A1', 'foo');
    expect(dn.getNames('blort!A1')).to.deep.equal(['foo']);
    expect(dn.getNames('blort!$A$1')).to.deep.equal(['foo']);

    dn.add('blort!$B$4', 'bar');
    expect(dn.getNames('blort!B4')).to.deep.equal(['bar']);
    expect(dn.getNames('blort!$B$4')).to.deep.equal(['bar']);

    dn.add("'blo rt'!$B$4", 'bar');
    expect(dn.getNames("'blo rt'!$B$4")).to.deep.equal(['bar']);
    dn.add("'blo ,!rt'!$B$4", 'bar');
    expect(dn.getNames("'blo ,!rt'!$B$4")).to.deep.equal(['bar']);
  });

  it('removes names for cells', () => {
    const dn = new DefinedNames();

    dn.add('blort!A1', 'foo');
    dn.add('blort!A1', 'bar');
    dn.remove('blort!A1', 'foo');

    expect(dn.getNames('blort!A1')).to.deep.equal(['bar']);
  });

  // get ranges example
  it('gets the right ranges for a name', () => {
    const dn = new DefinedNames();

    dn.add('blort!A1', 'vertical');
    dn.add('blort!A2', 'vertical');
    dn.add('blort!A3', 'vertical');

    dn.add('blort!C1', 'horizontal');
    dn.add('blort!D1', 'horizontal');
    dn.add('blort!E1', 'horizontal');

    dn.add('blort!C3', 'square');
    dn.add('blort!D3', 'square');
    dn.add('blort!C4', 'square');
    dn.add('blort!D4', 'square');

    dn.add('other!A1', 'single');

    expect(dn.getRanges('vertical')).to.deep.equal({
      name: 'vertical',
      ranges: ['blort!$A$1:$A$3'],
    });
    expect(dn.getRanges('horizontal')).to.deep.equal({
      name: 'horizontal',
      ranges: ['blort!$C$1:$E$1'],
    });
    expect(dn.getRanges('square')).to.deep.equal({
      name: 'square',
      ranges: ['blort!$C$3:$D$4'],
    });
    expect(dn.getRanges('single')).to.deep.equal({
      name: 'single',
      ranges: ['other!$A$1'],
    });
  });

  it('splices', () => {
    const dn = new DefinedNames();
    dn.add('vertical!A1', 'vertical');
    dn.add('vertical!A2', 'vertical');
    dn.add('vertical!A3', 'vertical');
    dn.add('vertical!A4', 'vertical');

    dn.add('horizontal!A1', 'horizontal');
    dn.add('horizontal!B1', 'horizontal');
    dn.add('horizontal!C1', 'horizontal');
    dn.add('horizontal!D1', 'horizontal');

    ['A', 'B', 'C', 'D'].forEach((col) => {
      [1, 2, 3, 4].forEach((row) => {
        dn.add(`square!${col}${row}`, 'square');
      });
    });

    dn.add('single!A1', 'singleA1');
    dn.add('single!D1', 'singleD1');
    dn.add('single!A4', 'singleA4');
    dn.add('single!D4', 'singleD4');

    dn.spliceRows('vertical', 2, 2, 1);
    dn.spliceColumns('horizontal', 2, 2, 1);
    dn.spliceRows('square', 2, 2, 1);
    dn.spliceColumns('square', 2, 2, 1);
    dn.spliceRows('single', 2, 2, 1);
    dn.spliceColumns('single', 2, 2, 1);

    expect(dn.getRanges('vertical')).to.deep.equal({
      name: 'vertical',
      ranges: ['vertical!$A$1', 'vertical!$A$3'],
    });
    expect(dn.getRanges('horizontal')).to.deep.equal({
      name: 'horizontal',
      ranges: ['horizontal!$A$1', 'horizontal!$C$1'],
    });
    expect(dn.getRanges('square')).to.deep.equal({
      name: 'square',
      ranges: ['square!$A$1', 'square!$C$1', 'square!$A$3', 'square!$C$3'],
    });
    expect(dn.getRanges('singleA1')).to.deep.equal({
      name: 'singleA1',
      ranges: ['single!$A$1'],
    });
    expect(dn.getRanges('singleD1')).to.deep.equal({
      name: 'singleD1',
      ranges: ['single!$C$1'],
    });
    expect(dn.getRanges('singleA4')).to.deep.equal({
      name: 'singleA4',
      ranges: ['single!$A$3'],
    });
    expect(dn.getRanges('singleD4')).to.deep.equal({
      name: 'singleD4',
      ranges: ['single!$C$3'],
    });
  });

  it('creates matrix from model', () => {
    const dn = new DefinedNames();

    dn.model = [];
    dn.add('blort!A1', 'bar');
    dn.remove('blort!A1', 'foo');

    expect(dn.getNames('blort!A1')).to.deep.equal(['bar']);
  });

  it('skips values with invalid range', () => {
    const dn = new DefinedNames();
    dn.model = [
      { name: 'eq', ranges: ['"="'] },
      { name: 'ref', ranges: ['#REF!'] },
      { name: 'single', ranges: ['Sheet3!$A$1'] },
      { name: 'range', ranges: ['Sheet3!$A$2:$F$2228'] },
    ];

    expect(dn.model).to.deep.equal([
      { name: 'single', ranges: ['Sheet3!$A$1'] },
      { name: 'range', ranges: ['Sheet3!$A$2:$F$2228'] },
    ]);
  });

  it('expands a range location into every cell of the range', () => {
    const dn = new DefinedNames();
    dn.add('S!A1:B2', 'sq');

    expect(dn.getRanges('sq')).to.deep.equal({ name: 'sq', ranges: ['S!$A$1:$B$2'] });
    ['S!A1', 'S!B1', 'S!A2', 'S!B2'].forEach((cell) => {
      expect(dn.getNames(cell)).to.deep.equal(['sq']);
    });
    expect(dn.getNames('S!C1')).to.deep.equal([]);
  });

  it('forEach visits every named cell with its name', () => {
    const dn = new DefinedNames();
    dn.add('S!A1:B2', 'sq');
    dn.add('S!D4', 'one');

    const visited = [];
    dn.forEach((name, cell) => visited.push(`${name}:${cell.address}`));

    expect(visited.sort()).to.deep.equal(['one:D4', 'sq:A1', 'sq:A2', 'sq:B1', 'sq:B2']);
  });

  it('removeAllNames drops every name from a cell', () => {
    const dn = new DefinedNames();
    dn.add('S!A1', 'first');
    dn.add('S!A1', 'second');
    dn.add('S!B1', 'first');

    dn.removeAllNames({ sheetName: 'S', address: 'A1', row: 1, col: 1 });

    expect(dn.getNames('S!A1')).to.deep.equal([]);
    expect(dn.getNames('S!B1')).to.deep.equal(['first']);
  });

  it('returns no ranges for an unknown name', () => {
    const dn = new DefinedNames();
    expect(dn.getRanges('nothing')).to.deep.equal({ name: 'nothing', ranges: [] });
  });
});
