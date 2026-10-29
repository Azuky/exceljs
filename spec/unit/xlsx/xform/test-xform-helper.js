const { PassThrough } = require('readable-stream');
const { cloneDeep, each } = require('../../../utils/under-dash');
const { expectKnownFailure } = require('../../../utils/known-bug');
const xmlTree = require('../../../utils/xml-tree');
const CompyXform = require('./compy-xform');

const parseSax = verquire('utils/parse-sax');
const XmlStream = verquire('utils/xml-stream');
const BooleanXform = verquire('xlsx/xform/simple/boolean-xform');

function getExpectation(expectation, name) {
  if (!expectation.hasOwnProperty(name)) {
    throw new Error(`Expectation missing required field: ${name}`);
  }
  return cloneDeep(expectation[name]);
}

// ===============================================================================================================
// provides boilerplate examples for the four transform steps: prepare, render,  parse and reconcile
//  prepare: model => preparedModel
//  render:  preparedModel => xml
//  parse:  xml => parsedModel
//  reconcile: parsedModel => reconciledModel

// An expectation may declare a known lib bug per test kind in `knownDeviations`:
//   knownDeviations: {
//     render: { reason: /pinned failure message/, neutralise: (root) => { ...undo the deviation... } },
//     parse: { reason: /.../, neutralise: (model) => model-with-the-deviation-undone },
//   }
// Leave `neutralise` out when nothing of the result survives the bug.
// `render` also covers `renderIn`, `parse` also covers `parseIn`. For such a kind two tests exist:
//  1. (only when `neutralise` is given) the normal test, run on the actual output after `neutralise` has undone exactly the
//     deviation (render: edits the parsed xml tree, see spec/utils/xml-tree.js; parse: returns the
//     model), so everything else stays strictly compared;
//  2. a known-bug test: the raw comparison must still fail with `reason`. It goes red when the
//     lib is fixed (then drop the declaration).
function deviationFor(expectation, kind) {
  const deviations = expectation.knownDeviations || {};
  if (deviations[kind]) return deviations[kind];
  if (kind === 'renderIn') return deviations.render;
  if (kind === 'parseIn') return deviations.parse;
  return undefined;
}

function itWithDeviation(expectation, kind, title, run) {
  const deviation = deviationFor(expectation, kind);
  if (!deviation) {
    it(title, () => run());
    return;
  }
  // without `neutralise` nothing of the result survives the bug: only the known-deviation test exists
  if (deviation.neutralise) it(title, () => run(deviation.neutralise));
  it(
    `${title} (known deviation)`,
    expectKnownFailure(() => run(), deviation.reason),
  );
}

// parses an xml string with the given xform; resolves with its model, rejects with parse errors
const parseXml = (xform, xml) =>
  new Promise((resolve, reject) => {
    const stream = new PassThrough();
    stream.write(xml);
    stream.end();
    xform.parse(parseSax(stream)).then(resolve).catch(reject);
  });

// deep equal whose failure message also carries both models, so a deviation `reason` can pin the difference
const expectModel = (actual, expected) => {
  try {
    expect(actual).to.deep.equal(expected);
  } catch (e) {
    e.message = `${e.message}\nactual: ${JSON.stringify(actual)}\nexpected: ${JSON.stringify(expected)}`;
    throw e;
  }
};

const neutraliseXml = (xml, neutralise) => (neutralise ? xmlTree.transform(xml, neutralise) : xml);

const its = {
  prepare(expectation) {
    it('Prepare Model', () =>
      new Promise((resolve) => {
        const model = getExpectation(expectation, 'initialModel');
        const result = getExpectation(expectation, 'preparedModel');

        const xform = expectation.create();
        xform.prepare(model, expectation.options);
        expect(cloneDeep(model, false)).to.deep.equal(result);
        resolve();
      }));
  },

  render(expectation) {
    itWithDeviation(expectation, 'render', 'Render to XML', (neutralise) => {
      const model = getExpectation(expectation, 'preparedModel');
      const result = getExpectation(expectation, 'xml');

      const xform = expectation.create();
      const xmlStream = new XmlStream();
      xform.render(xmlStream, model, 0);

      expect(neutraliseXml(xmlStream.xml, neutralise)).xml.to.equal(result);
    });
  },

  'prepare-render': function (expectation) {
    // when implementation details get in the way of testing the prepared result
    it('Prepare and Render to XML', () =>
      new Promise((resolve) => {
        const model = getExpectation(expectation, 'initialModel');
        const result = getExpectation(expectation, 'xml');

        const xform = expectation.create();
        const xmlStream = new XmlStream();

        xform.prepare(model, expectation.options);
        xform.render(xmlStream, model);

        expect(xmlStream.xml).xml.to.equal(result);
        resolve();
      }));
  },

  renderIn(expectation) {
    itWithDeviation(expectation, 'renderIn', 'Render in Composite to XML ', (neutralise) => {
      const model = {
        pre: true,
        child: getExpectation(expectation, 'preparedModel'),
        post: true,
      };
      const result = `<compy><pre/>${getExpectation(expectation, 'xml')}<post/></compy>`;

      const xform = new CompyXform({
        tag: 'compy',
        children: [
          {
            name: 'pre',
            xform: new BooleanXform({ tag: 'pre', attr: 'val' }),
          },
          { name: 'child', xform: expectation.create() },
          {
            name: 'post',
            xform: new BooleanXform({ tag: 'post', attr: 'val' }),
          },
        ],
      });

      const xmlStream = new XmlStream();
      xform.render(xmlStream, model);

      expect(neutraliseXml(xmlStream.xml, neutralise)).xml.to.equal(result);
    });
  },

  parseIn(expectation) {
    itWithDeviation(expectation, 'parseIn', 'Parse within composite', async (neutralise) => {
      const xml = `<compy><pre/>${getExpectation(expectation, 'xml')}<post/></compy>`;
      const childXform = expectation.create();
      const result = { pre: true };
      result[childXform.tag] = getExpectation(expectation, 'parsedModel');
      result.post = true;
      const xform = new CompyXform({
        tag: 'compy',
        children: [
          {
            name: 'pre',
            xform: new BooleanXform({ tag: 'pre', attr: 'val' }),
          },
          { name: childXform.tag, xform: childXform },
          {
            name: 'post',
            xform: new BooleanXform({ tag: 'post', attr: 'val' }),
          },
        ],
      });
      const model = await parseXml(xform, xml);
      // eliminate the undefined
      const clone = cloneDeep(model, false);
      if (neutralise) clone[childXform.tag] = neutralise(clone[childXform.tag]);
      expectModel(clone, result);
    });
  },

  parse(expectation) {
    itWithDeviation(expectation, 'parse', 'Parse to Model', async (neutralise) => {
      const xml = getExpectation(expectation, 'xml');
      const result = getExpectation(expectation, 'parsedModel');

      const model = await parseXml(expectation.create(), xml);
      // eliminate the undefined
      let clone = cloneDeep(model, false);
      if (neutralise) clone = neutralise(clone);
      expectModel(clone, result);
    });
  },

  reconcile(expectation) {
    it('Reconcile Model', () =>
      new Promise((resolve) => {
        const model = getExpectation(expectation, 'parsedModel');
        const result = getExpectation(expectation, 'reconciledModel');

        const xform = expectation.create();
        xform.reconcile(model, expectation.options);

        // eliminate the undefined
        const clone = cloneDeep(model, false);

        expect(clone).to.deep.equal(result);
        resolve();
      }));
  },
};

function testXform(expectations) {
  each(expectations, (expectation) => {
    const tests = getExpectation(expectation, 'tests');
    describe(expectation.title, () => {
      each(tests, (test) => {
        its[test](expectation);
      });
    });
  });
}

testXform.parseXml = parseXml;

module.exports = testXform;
