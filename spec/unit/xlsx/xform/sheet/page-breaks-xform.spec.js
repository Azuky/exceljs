const testXformHelper = require('../test-xform-helper');

const PageBreaksXform = verquire('xlsx/xform/sheet/page-breaks-xform');

const expectations = [
  {
    title: 'one page break',
    create() {
      return new PageBreaksXform();
    },
    initialModel: { id: 2, max: 3, min: 1, man: 1 },
    preparedModel: { id: 2, max: 3, min: 1, man: 1 },
    xml: '<brk id="2" max="3" min="1" man="1"/>',
    parsedModel: { id: 2, max: 3, min: 1, man: 1 },
    // known bug: parseOpen reads the (non-existent) `ref` attribute of <brk> instead of id/min/max/man
    knownFailures: ['parse'],
    knownFailureMessage: /expected undefined to deeply equal/,
    tests: ['prepare', 'render', 'renderIn', 'parse'],
  },
];

describe('PageBreaksXform', () => {
  testXformHelper(expectations);
});
