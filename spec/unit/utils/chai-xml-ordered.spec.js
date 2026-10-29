describe('chai-xml-ordered', () => {
  describe('xml.to.equal', () => {
    it('passes for identical documents', () => {
      expect('<a><b x="1">t</b><c/></a>').xml.to.equal('<a><b x="1">t</b><c/></a>');
    });

    it('ignores the XML declaration', () => {
      expect('<?xml version="1.0" encoding="UTF-8"?><a/>').xml.to.equal('<a/>');
    });

    it('ignores attribute order', () => {
      expect('<a x="1" y="2"/>').xml.to.equal('<a y="2" x="1"/>');
    });

    it('ignores whitespace between elements', () => {
      expect('<a>\n  <b/>\n  <c/>\n</a>').xml.to.equal('<a><b/><c/></a>');
    });

    it('treats CDATA as text', () => {
      expect('<a><![CDATA[x < y]]></a>').xml.to.equal('<a>x &lt; y</a>');
    });

    it('keeps namespace prefixes as part of names', () => {
      expect('<x:a xmlns:x="u"><x:b/></x:a>').xml.to.equal('<x:a xmlns:x="u"><x:b/></x:a>');
      expect('<x:a xmlns:x="u"/>').xml.not.to.equal('<y:a xmlns:y="u"/>');
    });

    it('fails when sibling elements are swapped', () => {
      expect(() => expect('<a><b/><c/></a>').xml.to.equal('<a><c/><b/></a>')).to.throw();
    });

    it('fails when repeated siblings are interleaved differently', () => {
      expect(() => expect('<a><b/><c/><b/></a>').xml.to.equal('<a><b/><b/><c/></a>')).to.throw();
    });

    it('fails on nested order differences', () => {
      expect(() => expect('<a><d><b/><c/></d></a>').xml.to.equal('<a><d><c/><b/></d></a>')).to.throw();
    });

    it('fails on different attribute values', () => {
      expect(() => expect('<a x="1"/>').xml.to.equal('<a x="2"/>')).to.throw();
    });

    it('fails on missing or extra attributes', () => {
      expect(() => expect('<a x="1"/>').xml.to.equal('<a/>')).to.throw();
      expect(() => expect('<a/>').xml.to.equal('<a x="1"/>')).to.throw();
    });

    it('fails on different text', () => {
      expect(() => expect('<a>one</a>').xml.to.equal('<a>two</a>')).to.throw();
    });

    it('fails on different tag names and child counts', () => {
      expect(() => expect('<a/>').xml.to.equal('<b/>')).to.throw();
      expect(() => expect('<a><b/></a>').xml.to.equal('<a><b/><b/></a>')).to.throw();
    });

    it('supports eq and equals aliases', () => {
      expect('<a><b/></a>').xml.to.eq('<a><b/></a>');
      expect('<a><b/></a>').xml.to.equals('<a><b/></a>');
      expect(() => expect('<a><b/><c/></a>').xml.to.eq('<a><c/><b/></a>')).to.throw();
    });

    it('treats empty documents as equal to each other only', () => {
      expect('').xml.to.equal('');
      expect(() => expect('').xml.to.equal('<a/>')).to.throw();
      expect(() => expect('<a/>').xml.to.equal('')).to.throw();
    });

    it('keeps a bare whitespace-only leaf as text, so <a> </a> differs from <a/>', () => {
      expect(() => expect('<a> </a>').xml.to.equal('<a/>')).to.throw();
      expect('<a> </a>').xml.to.equal('<a> </a>');
    });

    it('does not trim a bare whitespace-only leaf under .deep either', () => {
      expect(() => expect('<a> </a>').xml.deep.equal('<a/>')).to.throw();
      expect('<a> </a>').xml.deep.equal('<a> </a>');
    });

    it('compares text of mixed content as one concatenated chunk, ignoring child position', () => {
      expect('<a>x<b/>y</a>').xml.to.equal('<a>xy<b/></a>');
      expect(() => expect('<a>x<b/>y</a>').xml.to.equal('<a>x<b/>z</a>')).to.throw();
    });

    it('falls back to plain equal outside an xml chain', () => {
      expect('<a/>').to.equal('<a/>');
      expect(() => expect('<a/>').to.equal('<b/>')).to.throw();
    });
  });

  describe('xml.not.to.equal', () => {
    it('passes when sibling order differs', () => {
      expect('<a><b/><c/></a>').xml.not.to.equal('<a><c/><b/></a>');
    });

    it('fails when documents are equal', () => {
      expect(() => expect('<a><b/></a>').xml.not.to.equal('<a><b/></a>')).to.throw();
    });
  });

  describe('xml.to.be.valid', () => {
    it('passes for well-formed XML', () => {
      expect('<a><b/></a>').xml.to.be.valid();
    });

    it('fails for malformed XML', () => {
      expect(() => expect('<a><b></a>').xml.to.be.valid()).to.throw();
    });

    it('supports .not for malformed XML', () => {
      expect('<a><b></a>').xml.not.to.be.valid();
    });

    it('requires a string subject', () => {
      expect(() => expect(42).xml.to.be.valid()).to.throw();
    });
  });
});
