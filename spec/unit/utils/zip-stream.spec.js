const { Writable } = require('stream');
const JSZip = require('jszip');

const { ZipWriter } = verquire('utils/zip-stream');

describe('ZipWriter', () => {
  it('packs appended entries into a zip readable from the stream', async () => {
    const zip = new ZipWriter();
    zip.append('hello', { name: 'a.txt' });
    zip.append(Buffer.from('world'), { name: 'dir/b.txt' });
    await zip.finalize();

    const loaded = await JSZip.loadAsync(zip.read());
    expect(Object.keys(loaded.files).sort()).to.deep.equal(['a.txt', 'dir/', 'dir/b.txt']);
    expect(await loaded.file('a.txt').async('string')).to.equal('hello');
    expect(await loaded.file('dir/b.txt').async('string')).to.equal('world');
  });

  it('decodes base64 entries when the base64 option is set', async () => {
    const zip = new ZipWriter();
    zip.append(Buffer.from('decoded text').toString('base64'), { name: 'b64.txt', base64: true });
    zip.append('plain', { name: 'plain.txt', base64: false });
    await zip.finalize();

    const loaded = await JSZip.loadAsync(zip.read());
    expect(await loaded.file('b64.txt').async('string')).to.equal('decoded text');
    expect(await loaded.file('plain.txt').async('string')).to.equal('plain');
  });

  it('emits finish once the zip has been generated', async () => {
    const zip = new ZipWriter();
    const finished = new Promise((resolve) => zip.on('finish', resolve));
    zip.append('x', { name: 'x.txt' });
    await zip.finalize();
    await finished;
    expect(zip.read().length).to.be.greaterThan(0);
  });

  it('honours the compression option', async () => {
    const text = 'abc'.repeat(1000);
    const stored = new ZipWriter({ compression: 'STORE' });
    stored.append(text, { name: 'a.txt' });
    await stored.finalize();
    const deflated = new ZipWriter({ compression: 'DEFLATE' });
    deflated.append(text, { name: 'a.txt' });
    await deflated.finalize();

    expect(deflated.read().length).to.be.lessThan(stored.read().length);
  });

  it('pipes the finished zip into a destination', async () => {
    const zip = new ZipWriter();
    zip.append('piped', { name: 'p.txt' });
    await zip.finalize();

    const chunks = [];
    const done = new Promise((resolve, reject) => {
      const destination = new Writable({
        write(chunk, encoding, callback) {
          chunks.push(chunk);
          callback();
        },
      });
      destination.on('finish', resolve);
      destination.on('error', reject);
      zip.pipe(destination);
    });
    await done;

    const loaded = await JSZip.loadAsync(Buffer.concat(chunks));
    expect(await loaded.file('p.txt').async('string')).to.equal('piped');
  });

  it('reports the paused state set through pause and resume', () => {
    const zip = new ZipWriter();
    expect(zip.isPaused()).to.equal(false);
    zip.pause();
    expect(zip.isPaused()).to.equal(true);
    zip.resume();
    expect(zip.isPaused()).to.equal(false);
  });

  it('unpipe stops a destination from receiving data', async () => {
    const zip = new ZipWriter();
    zip.append('x', { name: 'x.txt' });
    const received = [];
    const destination = new Writable({
      write(chunk, encoding, callback) {
        received.push(chunk);
        callback();
      },
    });
    zip.pause();
    zip.pipe(destination);
    zip.unpipe(destination);
    await zip.finalize();
    expect(received).to.deep.equal([]);
  });

  it('does not implement unshift and wrap', () => {
    const zip = new ZipWriter();
    expect(() => zip.unshift(Buffer.from('x'))).to.throw('Not Implemented');
    expect(() => zip.wrap({})).to.throw('Not Implemented');
  });

  it('reads the first slice of the zip when given a size', async () => {
    const zip = new ZipWriter();
    zip.append('slice me', { name: 's.txt' });
    await zip.finalize();
    const whole = Buffer.from(zip.stream.toBuffer());
    const head = zip.read(10);
    expect(head.length).to.equal(10);
    expect(head.equals(whole.subarray(0, 10))).to.equal(true);
    // zip local file header signature
    expect(head.subarray(0, 4).toString('hex')).to.equal('504b0304');
  });

  it('setEncoding is forwarded to the underlying stream', () => {
    const zip = new ZipWriter();
    zip.setEncoding('base64');
    expect(zip.stream.encoding).to.equal('base64');
  });
});
