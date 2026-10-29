const { knownBug } = require('../../utils/known-bug');

const fs = require('fs');
const path = require('path');

const StreamBuf = verquire('utils/stream-buf');
const StringBuf = verquire('utils/string-buf');

describe('StreamBuf', () => {
  // StreamBuf is designed as a general-purpose writable-readable stream
  // However its use in ExcelJS is primarily as a memory buffer between
  // the streaming writers and the archive, hence the tests here will
  // focus just on that.
  it('writes strings as UTF8', () => {
    const stream = new StreamBuf();
    stream.write('Hello, World!');
    const chunk = stream.read();
    expect(chunk instanceof Buffer).to.be.ok;
    expect(chunk.toString('UTF8')).to.equal('Hello, World!');
  });

  it('writes StringBuf chunks', () => {
    const stream = new StreamBuf();
    const strBuf = new StringBuf({ size: 64 });
    strBuf.addText('Hello, World!');
    stream.write(strBuf);
    const chunk = stream.read();
    expect(chunk instanceof Buffer).to.be.ok;
    expect(chunk.toString('UTF8')).to.equal('Hello, World!');
  });

  it('signals end', () =>
    new Promise((resolve) => {
      const stream = new StreamBuf();
      stream.on('finish', () => {
        resolve();
      });
      stream.write('Hello, World!');
      stream.end();
    }));

  it('handles buffers', () =>
    new Promise((resolve, reject) => {
      const s = fs.createReadStream(path.join(__dirname, 'data/image1.png'));
      const sb = new StreamBuf();
      sb.on('finish', () => {
        const buf = sb.toBuffer();
        expect(buf.length).to.equal(1672);
        resolve();
      });
      sb.on('error', reject);
      s.pipe(sb);
    }));
  it('handle unsupported type of chunk', async () => {
    const stream = new StreamBuf();
    try {
      await stream.write({});
      expect.fail('should fail for given argument');
    } catch (e) {
      expect(e.message).to.equal('Chunk must be one of type String, Buffer or StringBuf.');
    }
  });

  describe('partial reads', () => {
    // known bug: BufferChunk read(size) copies up to index `size` instead of `iRead + size`, so a second partial read returns NUL bytes — see stage-2 notes
    knownBug(
      'reads a buffer back in consecutive slices',
      () => {
        const stream = new StreamBuf();
        stream.write('abcdefghij');
        expect(stream.read(3).toString()).to.equal('abc');
        expect(stream.read(3).toString()).to.equal('def');
        expect(stream.read(10).toString()).to.equal('ghij');
      },
      /expected '\\u0000\\u0000\\u0000' to equal 'def'/,
    );

    it('returns the first slice of a buffer for a partial read', () => {
      const stream = new StreamBuf();
      stream.write('abcdefghij');
      expect(stream.read(3).toString()).to.equal('abc');
    });

    it('returns everything when asked for more than is buffered', () => {
      const stream = new StreamBuf();
      stream.write('abcde');
      expect(stream.read(100).toString()).to.equal('abcde');
    });
  });

  describe('write after end', () => {
    // characterises current behaviour: unlike a plain node Writable, StreamBuf does not reject writes after end()
    it('keeps accepting writes after end() and appends them to the buffer', async () => {
      const stream = new StreamBuf();
      stream.write('ab');
      stream.end();
      const result = await stream.write('cd');
      expect(result).to.equal(true);
      expect(stream.toBuffer().toString()).to.equal('abcd');
    });
  });

  describe('error event', () => {
    it('delivers an emitted error to listeners and keeps buffered data readable', () => {
      const stream = new StreamBuf();
      const errors = [];
      stream.on('error', (e) => errors.push(e));
      stream.write('data');
      const failure = new Error('boom');
      stream.emit('error', failure);
      expect(errors).to.deep.equal([failure]);
      expect(stream.read().toString()).to.equal('data');
    });
  });
});
