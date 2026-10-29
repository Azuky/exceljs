const crypto = require('crypto');

const Encryptor = verquire('utils/encryptor');

describe('Encryptor', () => {
  // salt = bytes 0..15; expected values computed independently with python hashlib
  // (ECMA-376 password hash: H0 = H(salt + utf16le(pw)), Hn = H(Hn-1 + uint32le(n-1)))
  const salt = 'AAECAwQFBgcICQoLDA0ODw==';
  const password = 'pässword€';

  describe('hash', () => {
    it('hashes the concatenation of the given buffers', () => {
      const digest = Encryptor.hash('sha256', Buffer.from('a'), Buffer.from('bc'));
      expect(digest).to.be.an.instanceOf(Buffer);
      expect(digest.toString('hex')).to.equal('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
    });

    it('gives the same digest however the input is split', () => {
      const whole = Encryptor.hash('md5', Buffer.from('abcd'));
      const split = Encryptor.hash('md5', Buffer.from('ab'), Buffer.from('cd'));
      expect(split.toString('hex')).to.equal('e2fc714c4727ee9395f324cd2e7f331f');
      expect(split.equals(whole)).to.equal(true);
    });

    it('throws for an unknown algorithm', () => {
      expect(() => Encryptor.hash('nope-256', Buffer.from('a'))).to.throw();
    });
  });

  describe('convertPasswordToHash', () => {
    const cases = [
      ['sha512', 0, 'tjFY0DLYw12l8NWRNDNhV/pqGxfs3JRzUyG5Tv/f+PRG4y5L0CHqCOKow2vKcUquGyqfRW5YcVEFYp03BrRduQ=='],
      ['sha512', 1, 'qRmgVo0GpE4noyGEwwFeaFYa5KVOcp200fElkAfsoc9QgPLXLwyGb/DaRkGHAyqeo69thMkcY4L/BuDYwlBPTQ=='],
      ['sha512', 100, 'MRm25MJZ4WgDUwXuXNnNfUJfEGAmvgrhiY88eisoTQSAR1IbwbGqJmt8EmqdMmSzbm+pI18KeG8E404AIB1GoQ=='],
      ['sha1', 0, 'dPK4ylE10X2k+1+dCMeecRxeCB0='],
      ['sha1', 1, '4bXOD5JeMvWVE7aJ48YCs/+bzvo='],
      ['sha1', 100, '/dmWmHcJ1zNNaWvJVbQPl//xZVw='],
    ];
    cases.forEach(([algorithm, spinCount, expected]) => {
      it(`matches the known ${algorithm} hash with spinCount ${spinCount}`, () => {
        expect(Encryptor.convertPasswordToHash(password, algorithm, salt, spinCount)).to.equal(expected);
      });
    });

    it('matches a spin loop computed directly with node crypto', () => {
      let key = crypto
        .createHash('sha512')
        .update(Buffer.concat([Buffer.from(salt, 'base64'), Buffer.from('secret', 'utf16le')]))
        .digest();
      for (let i = 0; i < 25; i++) {
        const counter = Buffer.alloc(4);
        counter.writeUInt32LE(i);
        key = crypto.createHash('sha512').update(key).update(counter).digest();
      }
      expect(Encryptor.convertPasswordToHash('secret', 'sha512', salt, 25)).to.equal(key.toString('base64'));
    });

    it('accepts the algorithm name in any case', () => {
      expect(Encryptor.convertPasswordToHash(password, 'SHA1', salt, 1)).to.equal('4bXOD5JeMvWVE7aJ48YCs/+bzvo=');
    });

    it('gives different hashes for different passwords or salts', () => {
      const base = Encryptor.convertPasswordToHash('one', 'sha512', salt, 3);
      expect(Encryptor.convertPasswordToHash('two', 'sha512', salt, 3)).not.to.equal(base);
      expect(Encryptor.convertPasswordToHash('one', 'sha512', 'AQIDBA==', 3)).not.to.equal(base);
    });

    it('throws for an unsupported hash algorithm', () => {
      expect(() => Encryptor.convertPasswordToHash('pw', 'nope-256', salt, 1)).to.throw(
        "Hash algorithm 'nope-256' not supported!",
      );
    });
  });

  describe('randomBytes', () => {
    it('returns a buffer of the requested size', () => {
      [0, 1, 16, 64].forEach((size) => {
        const bytes = Encryptor.randomBytes(size);
        expect(bytes).to.be.an.instanceOf(Buffer);
        expect(bytes.length).to.equal(size);
      });
    });

    it('returns different bytes on each call', () => {
      expect(Encryptor.randomBytes(32).equals(Encryptor.randomBytes(32))).to.equal(false);
    });
  });
});
