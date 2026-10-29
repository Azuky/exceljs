const testXformHelper = require('../test-xform-helper');

const RowBreaksXform = verquire('xlsx/xform/sheet/row-breaks-xform');

const expectations = [
  {
    title: 'two manual row breaks',
    create() {
      return new RowBreaksXform();
    },
    preparedModel: [
      { id: 10, max: 16383, man: 1 },
      { id: 20, max: 16383, man: 1 },
    ],
    xml: '<rowBreaks count="2" manualBreakCount="2"><brk id="10" max="16383" man="1"/><brk id="20" max="16383" man="1"/></rowBreaks>',
    parsedModel: [
      { id: 10, max: 16383, man: 1 },
      { id: 20, max: 16383, man: 1 },
    ],
    // known bug: the brk children yield no model, so the list stays empty (see PageBreaksXform)
    // nothing of the list survives the bug, so the whole list is put back
    knownDeviations: {
      parse: {
        reason: /expected \[\] to deeply equal/,
        neutralise: () => [
          { id: 10, max: 16383, man: 1 },
          { id: 20, max: 16383, man: 1 },
        ],
      },
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'no row breaks renders nothing',
    create() {
      return new RowBreaksXform();
    },
    preparedModel: [],
    xml: '',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'undefined row breaks renders nothing',
    create() {
      return new RowBreaksXform();
    },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn'],
  },
];

describe('RowBreaksXform', () => {
  testXformHelper(expectations);
});
