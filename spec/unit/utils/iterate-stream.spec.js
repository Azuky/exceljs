const { Readable } = require('stream');
const ReadableStream = require('readable-stream').Readable;

const iterateStream = verquire('utils/iterate-stream');

async function collect(stream) {
  const items = [];
  for await (const item of iterateStream(stream)) {
    items.push(item);
  }
  return items;
}

describe('iterateStream', () => {
  it('yields every item of a normal object-mode stream in order', async () => {
    const items = await collect(Readable.from(['a', 'b', 'c', 'd']));
    expect(items).to.deep.equal(['a', 'b', 'c', 'd']);
  }, 2000);

  it('ignores a spurious manual end event and still yields all buffered items', async () => {
    const stream = new Readable({ objectMode: true, read() {} });
    const pushed = [1, 2, 3, 4, 5];
    stream.push(pushed[0]);

    const items = [];
    for await (const item of iterateStream(stream)) {
      items.push(item);
      if (items.length === 1) {
        // consumer is mid-loop, so iterateStream has paused the stream: these stay buffered
        // in the readable side instead of being emitted as 'data'
        pushed.slice(1).forEach((x) => stream.push(x));
        // spurious end (as unzipper's Parse does on writable 'finish')
        stream.emit('end');
        // genuine end arrives later
        setImmediate(() => stream.push(null));
      }
    }
    expect(items).to.deep.equal(pushed);
  }, 2000);

  it('terminates and yields data for a stream without readableEnded', async () => {
    const stream = new ReadableStream({ objectMode: true, read() {} });
    expect('readableEnded' in stream).to.equal(false);
    ['x', 'y', 'z'].forEach((item) => stream.push(item));
    stream.push(null);

    const items = await collect(stream);
    expect(items).to.deep.equal(['x', 'y', 'z']);
  }, 2000);
});
