const colCache = verquire('utils/col-cache');
const Cell = verquire('doc/cell');
const Enums = verquire('doc/enums');
const Excel = verquire('exceljs');

const sheetMock = {
  reset() {
    this.rows = [];
    this.columns = [];
  },
  findRow(num) {
    return this.rows[num];
  },
  getRow(num) {
    return this.rows[num] || this.createRow(num);
  },
  findColumn(num) {
    return this.columns[num];
  },
  getColumn(num) {
    return this.columns[num] || this.createColumn(num);
  },
  createRow(num) {
    this.rows[num] = {
      cells: [],
      findCell(col) {
        return this.cells[col];
      },
      getCell(col) {
        return this.cells[col] || this.createCell(col);
      },
      createCell(col) {
        const address = colCache.encodeAddress(this.number, col);
        const column = sheetMock.getColumn(col);
        return (this.cells[col] = new Cell(this, column, address));
      },
      number: num,
      get worksheet() {
        return sheetMock;
      },
    };
    return this.rows[num];
  },
  createColumn(num) {
    this.columns[num] = {
      number: num,
      letter: colCache.n2l(num),
    };
    return this.columns[num];
  },
  getCell(address) {
    const fullAddress = colCache.decodeAddress(address);
    const row = this.getRow(fullAddress.row);
    return row.getCell(fullAddress.col);
  },
  findCell(address) {
    const fullAddress = colCache.decodeAddress(address);
    const row = this.getRow(fullAddress.row);
    return row && row.findCell(fullAddress.col);
  },
};
const fonts = {
  arialBlackUI14: {
    name: 'Arial Black',
    family: 2,
    size: 14,
    underline: true,
    italic: true,
  },
  comicSansUdB16: {
    name: 'Comic Sans MS',
    family: 4,
    size: 16,
    underline: 'double',
    bold: true,
  },
  broadwayRedOutline20: {
    name: 'Broadway',
    family: 5,
    size: 20,
    outline: true,
    color: { argb: 'FFFF0000' },
  },
};

describe('Cell', () => {
  beforeEach(() => {
    sheetMock.reset();
  });
  it('stores values', () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell('A1');

    expect(a1.address).to.equal('A1');
    expect(a1.$col$row).to.equal('$A$1');

    expect(a1.type).to.equal(Enums.ValueType.Null);

    expect((a1.value = 5)).to.equal(5);
    expect(a1.value).to.equal(5);
    expect(a1.type).to.equal(Enums.ValueType.Number);

    const strValue = 'Hello, World!';
    expect((a1.value = strValue)).to.equal(strValue);
    expect(a1.value).to.equal(strValue);
    expect(a1.type).to.equal(Enums.ValueType.String);

    const dateValue = new Date();
    expect((a1.value = dateValue)).to.equal(dateValue);
    expect(a1.value).to.equal(dateValue);
    expect(a1.type).to.equal(Enums.ValueType.Date);

    let formulaValue = { formula: 'A2', result: 5 };
    expect((a1.value = formulaValue)).to.deep.equal(formulaValue);
    expect(a1.value).to.deep.equal(formulaValue);
    expect(a1.type).to.equal(Enums.ValueType.Formula);

    // no result
    formulaValue = { formula: 'A3' };
    expect((a1.value = formulaValue)).to.deep.equal(formulaValue);
    expect(a1.value).to.deep.equal({ formula: 'A3' });
    expect(a1.type).to.equal(Enums.ValueType.Formula);

    const hyperlinkValue = {
      hyperlink: 'http://www.link.com',
      text: 'www.link.com',
    };
    expect((a1.value = hyperlinkValue)).to.deep.equal(hyperlinkValue);
    expect(a1.value).to.deep.equal(hyperlinkValue);
    expect(a1.type).to.equal(Enums.ValueType.Hyperlink);

    expect((a1.value = null)).to.be.null;
    expect(a1.type).to.equal(Enums.ValueType.Null);

    expect((a1.value = { json: 'data' })).to.deep.equal({ json: 'data' });
    expect(a1.type).to.equal(Enums.ValueType.String);
  });
  it('validates options on construction', () => {
    const row = sheetMock.getRow(1);
    const column = sheetMock.getColumn(1);

    expect(() => {
      new Cell();
    }).to.throw(Error);
    expect(() => {
      new Cell(row);
    }).to.throw(Error);
    expect(() => {
      new Cell(row, 'A');
    }).to.throw(Error);
    expect(() => {
      new Cell(row, 'Hello, World!');
    }).to.throw(Error);
    expect(() => {
      new Cell(null, null, 'A1');
    }).to.throw(Error);
    expect(() => {
      new Cell(row, null, 'A1');
    }).to.throw(Error);
    expect(() => {
      new Cell(null, column, 'A1');
    }).to.throw(Error);
  });
  it('merges', () => {
    const a1 = sheetMock.getCell('A1');
    const a2 = sheetMock.getCell('A2');

    a1.value = 5;
    a2.value = 'Hello, World!';

    a2.merge(a1);

    expect(a2.value).to.equal(5);
    expect(a2.type).to.equal(Enums.ValueType.Merge);
    expect(a1._mergeCount).to.equal(1);
    expect(a1.isMerged).to.be.ok;
    expect(a2.isMerged).to.be.ok;
    expect(a2.isMergedTo(a1)).to.be.ok;
    expect(a2.master).to.equal(a1);
    expect(a1.master).to.equal(a1);

    // assignment of slaves write to the master
    a2.value = 7;
    expect(a1.value).to.equal(7);

    // assignment of strings should add 1 ref
    const strValue = 'Boo!';
    a2.value = strValue;
    expect(a1.value).to.equal(strValue);

    // unmerge should work also
    a2.unmerge();
    expect(a2.type).to.equal(Enums.ValueType.Null);
    expect(a1._mergeCount).to.equal(0);
    expect(a1.isMerged).to.not.be.ok;
    expect(a2.isMerged).to.not.be.ok;
    expect(a2.isMergedTo(a1)).to.not.be.ok;
    expect(a2.master).to.equal(a2);
    expect(a1.master).to.equal(a1);
  });

  it('upgrades from string to hyperlink', () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell('A1');

    const strValue = 'www.link.com';
    const linkValue = 'http://www.link.com';

    a1.value = strValue;

    a1._upgradeToHyperlink(linkValue);

    expect(a1.type).to.equal(Enums.ValueType.Hyperlink);
  });

  it("doesn't upgrade from non-string to hyperlink", () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell('A1');

    const linkValue = 'http://www.link.com';

    // null
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.equal(Enums.ValueType.Null);

    // number
    a1.value = 5;
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.equal(Enums.ValueType.Number);

    // date
    a1.value = new Date();
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.equal(Enums.ValueType.Date);

    // formula
    a1.value = { formula: 'A2' };
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.equal(Enums.ValueType.Formula);

    // hyperlink
    a1.value = { hyperlink: 'http://www.link2.com', text: 'www.link2.com' };
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.deep.equal(Enums.ValueType.Hyperlink);

    // cleanup
    a1.value = null;
  });

  it('inherits column styles', () => {
    sheetMock.getRow(1);
    const column = sheetMock.getColumn(1);

    column.style = {
      font: fonts.arialBlackUI14,
    };

    const a1 = sheetMock.getCell('A1');
    expect(a1.font).to.deep.equal(fonts.arialBlackUI14);
  });

  it('inherits row styles', () => {
    const row = sheetMock.getRow(1);
    sheetMock.getColumn(1);

    row.style = {
      font: fonts.broadwayRedOutline20,
    };

    const a1 = sheetMock.getCell('A1');
    expect(a1.font).to.deep.equal(fonts.broadwayRedOutline20);
  });

  it('has effective types', () => {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    const a1 = sheetMock.getCell('A1');

    expect(a1.type).to.equal(Enums.ValueType.Null);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Null);

    a1.value = 5;
    expect(a1.type).to.equal(Enums.ValueType.Number);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Number);

    a1.value = 'Hello, World!';
    expect(a1.type).to.equal(Enums.ValueType.String);
    expect(a1.effectiveType).to.equal(Enums.ValueType.String);

    a1.value = new Date();
    expect(a1.type).to.equal(Enums.ValueType.Date);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Date);

    a1.value = { formula: 'A2', result: 5 };
    expect(a1.type).to.deep.equal(Enums.ValueType.Formula);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Number);

    a1.value = { formula: 'A2', result: 'Hello, World!' };
    expect(a1.type).to.deep.equal(Enums.ValueType.Formula);
    expect(a1.effectiveType).to.equal(Enums.ValueType.String);

    a1.value = { hyperlink: 'http://www.link.com', text: 'www.link.com' };
    expect(a1.type).to.deep.equal(Enums.ValueType.Hyperlink);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Hyperlink);
  });

  it('shares formulas', () => {
    const a1 = sheetMock.getCell('A1');
    const b1 = sheetMock.getCell('B1');
    const c1 = sheetMock.getCell('C1');

    a1.value = 1;
    b1.value = { formula: 'A1+1', result: 2 };
    c1.value = { sharedFormula: 'B1', result: 3 };

    expect(b1.type).to.equal(Enums.ValueType.Formula);
    expect(b1.formulaType).to.equal(Enums.FormulaType.Master);
    expect(c1.type).to.equal(Enums.ValueType.Formula);
    expect(c1.formulaType).to.equal(Enums.FormulaType.Shared);
    expect(c1.formula).to.equal('B1+1');
  });

  it('escapes dangerous html', () => {
    const a1 = sheetMock.getCell('A1');

    a1.value = '<script>alert("yoohoo")</script>';

    expect(a1.html).to.equal('&lt;script&gt;alert(&quot;yoohoo&quot;)&lt;/script&gt;');
  });
  it('can set comment', () => {
    const a1 = sheetMock.getCell('A1');

    const comment = {
      texts: [
        {
          font: {
            size: 12,
            color: { theme: 0 },
            name: 'Calibri',
            family: 2,
            scheme: 'minor',
          },
          text: 'This is ',
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

    a1.note = comment;
    a1.value = 'test set value';

    expect(a1.model.comment.type).to.equal('note');
    expect(a1.model.comment.note).to.deep.equal(comment);
  });

  it('Cell comments supports setting margins, protection, and position properties', () => {
    const a1 = sheetMock.getCell('A1');

    const comment = {
      texts: [
        {
          font: {
            size: 12,
            color: { theme: 0 },
            name: 'Calibri',
            family: 2,
            scheme: 'minor',
          },
          text: 'This is ',
        },
      ],
      protection: {
        locked: 'False',
        lockText: 'True',
      },
    };

    a1.note = comment;
    a1.value = 'test set value';

    expect(a1.model.comment.type).to.equal('note');
    expect(a1.model.comment.note.texts).to.deep.equal(comment.texts);
    expect(a1.model.comment.note.protection).to.deep.equal(comment.protection);
    expect(a1.model.comment.note.margins.insetmode).to.equal('auto');
    expect(a1.model.comment.note.margins.inset).to.deep.equal([0.13, 0.13, 0.25, 0.25]);
    expect(a1.model.comment.note.editAs).to.equal('absolute');
  });

  describe('effective type of a formula', () => {
    function formulaCell(result) {
      const a1 = sheetMock.getCell('A1');
      a1.value = { formula: '1+1', result };
      return a1;
    }

    it('follows the type of a numeric, string or date result', () => {
      expect(formulaCell(2).effectiveType).to.equal(Enums.ValueType.Number);
      expect(formulaCell('s').effectiveType).to.equal(Enums.ValueType.String);
      expect(formulaCell(new Date('2020-01-02T03:04:05Z')).effectiveType).to.equal(Enums.ValueType.Date);
    });

    it('is a hyperlink for a hyperlink result', () => {
      const cell = formulaCell({ text: 'link', hyperlink: 'http://example.com' });
      expect(cell.type).to.equal(Enums.ValueType.Formula);
      expect(cell.effectiveType).to.equal(Enums.ValueType.Hyperlink);
    });

    it('is Null when there is no result', () => {
      expect(formulaCell(undefined).effectiveType).to.equal(Enums.ValueType.Null);
      expect(formulaCell(null).effectiveType).to.equal(Enums.ValueType.Null);
    });

    // known bug: FormulaValue.effectiveType has no branch for boolean results and reports Null — see stage-2 notes
    it.fails('is Boolean for a boolean result', () => {
      expect(formulaCell(true).effectiveType).to.equal(Enums.ValueType.Boolean);
    });

    // known bug: FormulaValue.effectiveType has no branch for error results and reports Null — see stage-2 notes
    it.fails('is Error for an error result', () => {
      expect(formulaCell({ error: '#N/A' }).effectiveType).to.equal(Enums.ValueType.Error);
    });

    it('exposes the formula, result and formula type', () => {
      const cell = formulaCell(2);
      expect(cell.formula).to.equal('1+1');
      expect(cell.result).to.equal(2);
      expect(cell.formulaType).to.equal(Enums.FormulaType.Master);
    });
  });

  describe('text and csv representations', () => {
    function cellWith(value) {
      const a1 = sheetMock.getCell('A1');
      a1.value = value;
      return a1;
    }

    it('toString renders the text of each value type', () => {
      expect(cellWith(null).toString()).to.equal('');
      expect(cellWith(5).toString()).to.equal('5');
      expect(cellWith('plain').toString()).to.equal('plain');
      expect(cellWith(true).toString()).to.equal('true');
      expect(cellWith({ error: '#N/A' }).toString()).to.equal('#N/A');
      expect(cellWith({ text: 'link', hyperlink: 'http://example.com' }).toString()).to.equal('link');
      expect(cellWith({ richText: [{ text: 'Hel' }, { text: 'lo' }] }).toString()).to.equal('Hello');
      expect(cellWith({ formula: '1+1', result: 2 }).toString()).to.equal('2');
    });

    it('toCsvString renders each value type as a csv field', () => {
      expect(cellWith(null).toCsvString()).to.equal('');
      expect(String(cellWith(5).toCsvString())).to.equal('5');
      expect(cellWith('say "hi"').toCsvString()).to.equal('"say ""hi"""');
      expect(String(cellWith(true).toCsvString())).to.equal('1');
      expect(cellWith({ error: '#N/A' }).toCsvString()).to.equal('#N/A');
      expect(cellWith({ text: 'link', hyperlink: 'http://example.com' }).toCsvString()).to.equal('http://example.com');
      expect(String(cellWith({ formula: '1+1', result: 3 }).toCsvString())).to.equal('3');
    });

    // known bug: RichTextValue.toCsvString reads this.text, which the value class does not have, and throws — see stage-2 notes
    it.fails('toCsvString renders a rich text cell as its quoted plain text', () => {
      expect(cellWith({ richText: [{ text: 'Hel' }, { text: 'lo' }] }).toCsvString()).to.equal('"Hello"');
    });

    it('isHyperlink is true only for hyperlink cells', () => {
      expect(cellWith({ text: 'link', hyperlink: 'http://example.com' }).isHyperlink).to.equal(true);
      expect(cellWith('http://example.com').isHyperlink).to.equal(false);
      expect(cellWith(5).isHyperlink).to.equal(false);
    });
  });

  describe('protection, names and destroy', () => {
    it('protection reads and writes the style', () => {
      const a1 = sheetMock.getCell('A1');
      expect(a1.protection).to.equal(undefined);
      a1.protection = { locked: false, hidden: true };
      expect(a1.protection).to.deep.equal({ locked: false, hidden: true });
      expect(a1.style.protection).to.deep.equal({ locked: false, hidden: true });
    });

    it('removeAllNames clears every name of the cell only', () => {
      const ws = new Excel.Workbook().addWorksheet('names');
      const a1 = ws.getCell('A1');
      const b1 = ws.getCell('B1');
      a1.names = ['first', 'second'];
      b1.name = 'first';
      expect(a1.names).to.have.members(['first', 'second']);

      a1.removeAllNames();

      expect(a1.names).to.deep.equal([]);
      expect(a1.name).to.equal(undefined);
      expect(b1.names).to.deep.equal(['first']);
    });

    it('addName and removeName adjust single names', () => {
      const ws = new Excel.Workbook().addWorksheet('names');
      const a1 = ws.getCell('A1');
      a1.addName('one');
      a1.addName('two');
      a1.removeName('one');
      expect(a1.names).to.deep.equal(['two']);
    });

    it('destroy releases the references held by the cell', () => {
      const ws = new Excel.Workbook().addWorksheet('names');
      const a1 = ws.getCell('A1');
      a1.value = 5;
      a1.destroy();
      expect(a1._value).to.equal(undefined);
      expect(a1._row).to.equal(undefined);
      expect(a1._column).to.equal(undefined);
      expect(a1.style).to.equal(undefined);
    });
  });
});
