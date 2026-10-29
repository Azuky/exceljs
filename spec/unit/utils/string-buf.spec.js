const StringBuf = verquire('utils/string-buf');

describe('StringBuf', () => {
  // StringBuf is a lightweight string-builder used by the streaming writers to build
  // strings (e.g. for row data) without too many memory operations
  it('writes strings as UTF8', () => {
    const sb = new StringBuf({ size: 64 });
    sb.addText('Hello, World!');
    const chunk = sb.toBuffer();
    expect(chunk.toString('UTF8')).to.equal('Hello, World!');
  });

  it('grows properly', () => {
    const sb = new StringBuf({ size: 8 });
    expect(sb.length).to.equal(0);
    expect(sb.capacity).to.equal(8);

    // write simple UTF8 string. Should use 7 bytes
    // that's within 4 bytes of 16
    sb.addText('Hello, ');
    expect(sb.length).to.equal(7);
    expect(sb.capacity).to.equal(16);

    // add some more (6 bytes)
    sb.addText('World!');
    expect(sb.length).to.equal(13);
    expect(sb.capacity).to.equal(32);

    // and more (7 bytes)
    sb.addText(' Hello.');
    expect(sb.length).to.equal(20);
    expect(sb.capacity).to.equal(32);

    // after all that - the string should be intact
    const chunk = sb.toBuffer();
    expect(chunk.toString('UTF8')).to.equal('Hello, World! Hello.');
  });

  it('resets', () => {
    const sb = new StringBuf({ size: 64 });
    sb.addText('Hello, ');
    expect(sb.length).to.equal(7);

    sb.reset();
    expect(sb.length).to.equal(0);

    sb.addText('World!');
    expect(sb.length).to.equal(6);

    const chunk = sb.toBuffer();
    expect(chunk.toString('UTF8')).to.equal('World!');
  });

  describe('addStringBuf', () => {
    it('appends the content of another StringBuf', () => {
      const sb = new StringBuf({ size: 64 });
      const other = new StringBuf({ size: 64 });
      sb.addText('Hello, ');
      other.addText('World!');
      sb.addStringBuf(other);
      expect(sb.toBuffer().toString('UTF8')).to.equal('Hello, World!');
      expect(sb.length).to.equal(13);
    });

    it('grows when the other buffer does not fit', () => {
      const sb = new StringBuf({ size: 8 });
      const other = new StringBuf({ size: 64 });
      sb.addText('abc');
      other.addText('0123456789012345678901234567890123456789');
      sb.addStringBuf(other);
      expect(sb.capacity).to.be.at.least(43);
      expect(sb.toBuffer().toString('UTF8')).to.equal('abc0123456789012345678901234567890123456789');
    });

    it('ignores an empty StringBuf and keeps the cached buffer valid', () => {
      const sb = new StringBuf({ size: 16 });
      sb.addText('abc');
      expect(sb.toBuffer().toString()).to.equal('abc');
      sb.addStringBuf(new StringBuf({ size: 16 }));
      expect(sb.length).to.equal(3);
      expect(sb.toBuffer().toString()).to.equal('abc');
    });

    it('invalidates a previously taken buffer when content is appended', () => {
      const sb = new StringBuf({ size: 16 });
      const other = new StringBuf({ size: 16 });
      sb.addText('ab');
      other.addText('cd');
      expect(sb.toBuffer().toString()).to.equal('ab');
      sb.addStringBuf(other);
      expect(sb.toBuffer().toString()).to.equal('abcd');
    });
  });
});
