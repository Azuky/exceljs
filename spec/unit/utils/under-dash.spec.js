const _ = verquire('utils/under-dash');
const util = require('util');

describe('under-dash', () => {
  describe('isEqual', () => {
    const values = [
      0,
      1,
      true,
      false,
      'string',
      'foobar',
      'other string',
      [],
      ['array'],
      ['array', 'foobar'],
      ['array2'],
      ['array2', 'foobar'],
      {},
      { object: 1 },
      { object: 2 },
      { object: 1, foobar: 'quux' },
      { object: 2, foobar: 'quux' },
      null,
      undefined,
      () => {},
      () => {},
      Symbol('foo'),
      Symbol('foo'),
      Symbol('bar'),
    ];

    function showVal(o) {
      return util.inspect(o, { compact: true });
    }

    it('works on simple values', () => {
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values.length; j++) {
          const a = values[i];
          const b = values[j];

          const assertion = `${showVal(a)} ${i === j ? '==' : '!='} ${showVal(b)}`;
          expect(_.isEqual(a, b)).to.equal(i === j, `expected ${assertion}`);
        }
      }
    });

    it('works on complex arrays', () => {
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values.length; j++) {
          const a = [values[i]];
          const b = [values[j]];

          const assertion = `${showVal(a)} ${i === j ? '==' : '!='} ${showVal(b)}`;
          expect(_.isEqual(a, b)).to.equal(i === j, `expected ${assertion}`);
        }
      }
    });

    it('works on complex objects', () => {
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values.length; j++) {
          const a = { key: values[i] };
          const b = { key: values[j] };

          const assertion = `${showVal(a)} ${i === j ? '==' : '!='} ${showVal(b)}`;
          expect(_.isEqual(a, b)).to.equal(i === j, `expected ${assertion}`);
        }
      }
    });
  });

  describe('collection helpers', () => {
    it('each iterates arrays with (value, index) and objects with (value, key)', () => {
      const seen = [];
      _.each([10, 20], (v, i) => seen.push([v, i]));
      _.each({ a: 1, b: 2 }, (v, k) => seen.push([v, k]));
      expect(seen).to.deep.equal([
        [10, 0],
        [20, 1],
        [1, 'a'],
        [2, 'b'],
      ]);
    });

    it('each tolerates null and undefined', () => {
      let calls = 0;
      _.each(null, () => calls++);
      _.each(undefined, () => calls++);
      expect(calls).to.equal(0);
    });

    it('some works on arrays, plain objects and null', () => {
      expect(_.some([1, 2, 3], (v) => v > 2)).to.equal(true);
      expect(_.some([1, 2, 3], (v) => v > 5)).to.equal(false);
      expect(_.some({ a: 1, b: 2 }, (v, k) => k === 'b' && v === 2)).to.equal(true);
      expect(_.some({ a: 1 }, (v) => v > 1)).to.equal(false);
      expect(_.some(null, () => true)).to.equal(false);
    });

    it('every works on arrays, plain objects and null', () => {
      expect(_.every([1, 2, 3], (v) => v > 0)).to.equal(true);
      expect(_.every([1, 2, 3], (v) => v > 1)).to.equal(false);
      expect(_.every({ a: 1, b: 2 }, (v) => v > 0)).to.equal(true);
      expect(_.every({ a: 1, b: 2 }, (v) => v > 1)).to.equal(false);
      expect(_.every(null, () => false)).to.equal(true);
    });

    it('map works on arrays, plain objects and null', () => {
      expect(_.map([1, 2], (v) => v * 2)).to.deep.equal([2, 4]);
      expect(_.map({ a: 1, b: 2 }, (v, k) => k + v)).to.deep.equal(['a1', 'b2']);
      expect(_.map(undefined, () => 1)).to.deep.equal([]);
    });

    it('keyBy indexes items by the given property', () => {
      expect(
        _.keyBy(
          [
            { id: 'x', n: 1 },
            { id: 'y', n: 2 },
          ],
          'id',
        ),
      ).to.deep.equal({
        x: { id: 'x', n: 1 },
        y: { id: 'y', n: 2 },
      });
    });
  });

  describe('escapeHtml', () => {
    it('escapes the five special characters', () => {
      expect(_.escapeHtml('<a href="x">Tom & \'Jerry\'</a>')).to.equal(
        '&lt;a href=&quot;x&quot;&gt;Tom &amp; &apos;Jerry&apos;&lt;/a&gt;',
      );
    });

    it('returns text without special characters unchanged', () => {
      expect(_.escapeHtml('plain text')).to.equal('plain text');
    });
  });

  describe('strcmp', () => {
    it('orders strings like a comparator', () => {
      expect(_.strcmp('a', 'b')).to.equal(-1);
      expect(_.strcmp('b', 'a')).to.equal(1);
      expect(_.strcmp('a', 'a')).to.equal(0);
    });
  });

  describe('isUndefined / isObject', () => {
    it('detects undefined', () => {
      expect(_.isUndefined(undefined)).to.equal(true);
      expect(_.isUndefined(null)).to.equal(false);
      expect(_.isUndefined(0)).to.equal(false);
    });

    it('detects plain objects only', () => {
      expect(_.isObject({})).to.equal(true);
      expect(_.isObject([])).to.equal(false);
      expect(_.isObject(null)).to.equal(false);
      expect(_.isObject('x')).to.equal(false);
    });
  });

  describe('deepMerge', () => {
    it('merges nested objects into the first argument', () => {
      const target = { a: { x: 1, y: 2 }, b: 1 };
      const result = _.deepMerge(target, { a: { y: 3, z: 4 }, c: 5 });
      expect(result).to.equal(target);
      expect(result).to.deep.equal({ a: { x: 1, y: 3, z: 4 }, b: 1, c: 5 });
    });

    it('merges arrays by index and ignores undefined values', () => {
      expect(_.deepMerge({ list: [1, 2, 3], k: 'keep' }, { list: [9], k: undefined })).to.deep.equal({
        list: [9, 2, 3],
        k: 'keep',
      });
    });

    it('merges several sources left to right and copies nested values instead of aliasing them', () => {
      const nested = { deep: true };
      const result = _.deepMerge({}, { a: 1 }, { a: 2, n: nested });
      expect(result).to.deep.equal({ a: 2, n: { deep: true } });
      expect(result.n).not.to.equal(nested);
    });
  });

  describe('isEqual with null and mixed types', () => {
    it('distinguishes null from objects and arrays in either position', () => {
      expect(_.isEqual(null, {})).to.equal(false);
      expect(_.isEqual({}, null)).to.equal(false);
      expect(_.isEqual(null, null)).to.equal(true);
      expect(_.isEqual([], {})).to.equal(false);
      expect(_.isEqual({}, [])).to.equal(false);
    });
  });
});
