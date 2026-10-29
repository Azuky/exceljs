const Excel = verquire('exceljs');
const Encryptor = verquire('utils/encryptor');

describe('Worksheet', () => {
  describe('protection', () => {
    function createSheet() {
      return new Excel.Workbook().addWorksheet('blort');
    }

    it('is not protected by default', () => {
      expect(createSheet().sheetProtection).to.not.be.ok;
    });

    it('protects with a password using SHA-512 and a verifiable hash', async () => {
      const ws = createSheet();
      await ws.protect('pw');
      const { sheetProtection } = ws;
      expect(sheetProtection.sheet).to.equal(true);
      expect(sheetProtection.algorithmName).to.equal('SHA-512');
      expect(sheetProtection.spinCount).to.equal(100000);
      expect(Buffer.from(sheetProtection.saltValue, 'base64')).to.have.length(16);
      expect(sheetProtection.hashValue).to.equal(
        Encryptor.convertPasswordToHash('pw', 'SHA512', sheetProtection.saltValue, 100000),
      );
    });

    it('uses a fresh salt for every call', async () => {
      const ws = createSheet();
      await ws.protect('pw', { spinCount: 1 });
      const first = ws.sheetProtection.saltValue;
      await ws.protect('pw', { spinCount: 1 });
      expect(ws.sheetProtection.saltValue).not.to.equal(first);
    });

    it('protects without a password using only the sheet flag', async () => {
      const ws = createSheet();
      await ws.protect();
      expect(ws.sheetProtection).to.deep.equal({ sheet: true });
    });

    it('clamps a negative spinCount to 0', async () => {
      const ws = createSheet();
      await ws.protect('pw', { spinCount: -5 });
      expect(ws.sheetProtection.spinCount).to.equal(0);
      expect(ws.sheetProtection.algorithmName).to.equal('SHA-512');
      expect(ws.sheetProtection.hashValue).to.be.a('string');
      expect(ws.sheetProtection.saltValue).to.be.a('string');
    });

    it('rounds a fractional spinCount', async () => {
      const ws = createSheet();
      await ws.protect('pw', { spinCount: 10.4 });
      expect(ws.sheetProtection.spinCount).to.equal(10);
    });

    it('falls back to 100000 for a non-finite spinCount', async () => {
      const ws = createSheet();
      await ws.protect('pw', { spinCount: Infinity });
      expect(ws.sheetProtection.spinCount).to.equal(100000);
      await ws.protect('pw', { spinCount: NaN });
      expect(ws.sheetProtection.spinCount).to.equal(100000);
    });

    it('merges options into the protection and drops spinCount when there is no password', async () => {
      const ws = createSheet();
      await ws.protect(null, { spinCount: 10, formatCells: false });
      expect(ws.sheetProtection).to.deep.equal({ sheet: true, formatCells: false });
      expect(ws.sheetProtection).not.to.have.property('spinCount');
    });

    it('lets options override the defaults when a password is given', async () => {
      const ws = createSheet();
      await ws.protect('pw', { spinCount: 3, sheet: false, objects: true });
      expect(ws.sheetProtection.sheet).to.equal(false);
      expect(ws.sheetProtection.objects).to.equal(true);
      expect(ws.sheetProtection.spinCount).to.equal(3);
    });

    it('unprotect clears the protection', async () => {
      const ws = createSheet();
      await ws.protect('pw', { spinCount: 1 });
      ws.unprotect();
      expect(ws.sheetProtection).to.equal(null);
    });

    it('unprotect on an unprotected sheet leaves it unprotected', () => {
      const ws = createSheet();
      ws.unprotect();
      expect(ws.sheetProtection).to.equal(null);
    });
  });
});
