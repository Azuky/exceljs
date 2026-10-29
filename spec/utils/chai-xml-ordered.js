/**
 * chai plugin: XML assertions where sibling element order is significant.
 *
 * Mirrors chai-xml 0.3.2 (`expect(str).xml.to.equal(str)`, `.xml.to.be.valid()`),
 * which compares via xml2js. Normalisation kept from it:
 *  - XML declaration, comments and processing instructions are ignored
 *  - attribute order is ignored
 *  - namespaces are not resolved: prefixes and xmlns:* attributes are plain names/attributes
 *  - whitespace-only text is dropped (unless CDATA, or the element has no attrs/children)
 *  - text is not trimmed, except in a `.deep` chain (chai-xml passes trim: flag('deep'))
 *  - text chunks and CDATA of one element are concatenated
 * The one difference: xml2js groups children by tag name, so it cannot see the order
 * of siblings. Here children are compared as an ordered list.
 */
const sax = require('sax');

function parse(xml) {
  // xml2js accepts an empty/blank document and yields null
  if (/^\s*$/.test(xml)) return null;
  const parser = sax.parser(true, { xmlns: false });
  const root = { children: [] };
  const stack = [root];
  let cdata = false;

  parser.onerror = (err) => {
    throw err;
  };
  parser.onopentag = (tag) => {
    const node = {
      name: tag.name,
      attrs: tag.attributes,
      text: '',
      cdata: false,
      children: [],
    };
    stack[stack.length - 1].children.push(node);
    stack.push(node);
  };
  parser.ontext = (text) => {
    stack[stack.length - 1].text = (stack[stack.length - 1].text || '') + text;
  };
  parser.onopencdata = () => {
    cdata = true;
  };
  parser.oncdata = (text) => {
    const node = stack[stack.length - 1];
    node.text = (node.text || '') + text;
    node.cdata = true;
  };
  parser.onclosecdata = () => {
    cdata = false;
  };
  parser.onclosetag = () => {
    stack.pop();
  };

  parser.write(xml).close();
  if (cdata || stack.length !== 1 || root.children.length !== 1) {
    throw new Error('Invalid XML: expected exactly one complete root element');
  }
  return root.children[0];
}

function normalizeText(node, trim) {
  let { text } = node;
  if (!node.cdata && /^\s*$/.test(text)) {
    // xml2js drops whitespace-only text, but keeps it for a bare leaf element
    const bare = !node.children.length && !Object.keys(node.attrs).length;
    if (!bare) text = '';
  }
  return trim ? text.trim() : text;
}

// Returns a description of the first difference, or null when equal.
function diff(a, b, path, trim) {
  if (a.name !== b.name) return `${path}: tag <${a.name}> !== <${b.name}>`;
  const here = path;
  const ak = Object.keys(a.attrs).sort();
  const bk = Object.keys(b.attrs).sort();
  if (ak.join('\u0000') !== bk.join('\u0000')) {
    return `${here}: attributes [${ak}] !== [${bk}]`;
  }
  for (const k of ak) {
    if (a.attrs[k] !== b.attrs[k]) {
      return `${here}@${k}: ${JSON.stringify(a.attrs[k])} !== ${JSON.stringify(b.attrs[k])}`;
    }
  }
  const at = normalizeText(a, trim);
  const bt = normalizeText(b, trim);
  if (at !== bt) return `${here}: text ${JSON.stringify(at)} !== ${JSON.stringify(bt)}`;
  if (a.children.length !== b.children.length) {
    const names = (n) => n.children.map((c) => c.name).join(',');
    return `${here}: children [${names(a)}] !== [${names(b)}]`;
  }
  for (let i = 0; i < a.children.length; i++) {
    const child = `${here}/${a.children[i].name}[${i}]`;
    if (a.children[i].name !== b.children[i].name) {
      const names = (n) => n.children.map((c) => c.name).join(',');
      return `${here}: child #${i} <${a.children[i].name}> !== <${b.children[i].name}> (order matters); actual [${names(a)}] vs expected [${names(b)}]`;
    }
    const d = diff(a.children[i], b.children[i], child, trim);
    if (d) return d;
  }
  return null;
}

module.exports = function chaiXmlOrdered(chai, utils) {
  const { flag } = utils;
  const { Assertion } = chai;

  Assertion.addProperty('xml', function () {
    new Assertion(this._obj).to.be.a('string');
    flag(this, 'xml', true);
  });

  Assertion.addMethod('valid', function () {
    new Assertion(flag(this, 'xml')).to.be.true;
    let err = null;
    try {
      parse(this._obj);
    } catch (e) {
      err = e;
    }
    this.assert(err === null, 'expected #{this} to be valid', 'expected #{this} not be not valid', err);
  });

  const compareXml = (_super) =>
    function assertEqual(value) {
      if (!flag(this, 'xml')) {
        return _super.apply(this, arguments);
      }
      new Assertion(value).to.be.a('string');
      const actual = parse(this._obj);
      const expected = parse(value);
      let d = null;
      if (actual === null || expected === null) {
        if (actual !== expected) d = 'one document is empty and the other is not';
      } else {
        d = diff(actual, expected, `/${actual.name}`, !!flag(this, 'deep'));
      }
      this.assert(
        d === null,
        `expected XML to equal (element order significant): ${d}`,
        'expected XML not to equal the given XML',
      );
      return undefined;
    };

  Assertion.overwriteMethod('equal', compareXml);
  Assertion.overwriteMethod('equals', compareXml);
  Assertion.overwriteMethod('eq', compareXml);
};
