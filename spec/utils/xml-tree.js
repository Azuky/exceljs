// Minimal XML tree for specs that need to undo ONE known deviation of a rendered document
// before comparing it with the expected one (see `knownDeviations` in test-xform-helper.js).
// parse() -> { name, attrs, text, children }; serialize() writes it back. Declarations and
// comments are dropped, which the XML comparison ignores anyway.
const sax = require('sax');

function parse(xml) {
  const parser = sax.parser(true, { xmlns: false });
  const root = { children: [] };
  const stack = [root];
  parser.onerror = (err) => {
    throw err;
  };
  parser.onopentag = (tag) => {
    const node = { name: tag.name, attrs: { ...tag.attributes }, text: '', children: [] };
    stack[stack.length - 1].children.push(node);
    stack.push(node);
  };
  parser.ontext = (text) => {
    stack[stack.length - 1].text += text;
  };
  parser.oncdata = (text) => {
    stack[stack.length - 1].text += text;
  };
  parser.onclosetag = () => {
    stack.pop();
  };
  parser.write(xml).close();
  return root.children[0];
}

const escapeText = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (s) => escapeText(s).replace(/"/g, '&quot;');

function serialize(node) {
  const attrs = Object.keys(node.attrs)
    .map((key) => ` ${key}="${escapeAttr(String(node.attrs[key]))}"`)
    .join('');
  const inner = escapeText(node.text) + node.children.map(serialize).join('');
  return inner ? `<${node.name}${attrs}>${inner}</${node.name}>` : `<${node.name}${attrs}/>`;
}

// every element called `name`, depth first, in document order
function findAll(node, name, found = []) {
  if (node.name === name) found.push(node);
  node.children.forEach((child) => findAll(child, name, found));
  return found;
}

// applies `neutralise(root)` to the parsed document and returns the xml again
function transform(xml, neutralise) {
  if (/^\s*$/.test(xml)) return xml;
  const root = parse(xml);
  neutralise(root);
  return serialize(root);
}

module.exports = { parse, serialize, findAll, transform };
