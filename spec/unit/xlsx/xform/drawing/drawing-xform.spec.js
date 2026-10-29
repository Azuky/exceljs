const fs = require('fs');

const testXformHelper = require('../test-xform-helper');
const { findAll } = require('../../../../utils/xml-tree');

// CT_OneCellAnchor has no attributes: undo only the editAs the lib writes there
const removeEditAs = (root) => {
  findAll(root, 'xdr:oneCellAnchor').forEach((anchor) => {
    delete anchor.attrs.editAs;
  });
};

const DrawingXform = verquire('xlsx/xform/drawing/drawing-xform');

const options = {
  rels: {
    rId1: { Target: '../media/image1.jpg' },
    rId2: { Target: '../media/image2.jpg' },
  },
  mediaIndex: { image1: 0, image2: 1 },
  media: [{}, {}],
};

const expectations = [
  {
    title: 'Drawing 1',
    create() {
      return new DrawingXform({ tag: 'xdr:from' });
    },
    initialModel: require('./data/drawing.1.0.js'),
    preparedModel: require('./data/drawing.1.1.js'),
    xml: fs.readFileSync(`${__dirname}/data/drawing.1.2.xml`).toString(),
    parsedModel: require('./data/drawing.1.3.js'),
    reconciledModel: require('./data/drawing.1.4.js'),
    // known bug: the oneCellAnchor is written with editAs (CT_OneCellAnchor has no attributes)
    knownDeviations: {
      render: {
        reason: /xdr:oneCellAnchor(\[\d+\])?: attributes \[editAs\] !== \[\]/,
        neutralise: removeEditAs,
      },
    },
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options,
  },
];

describe('DrawingXform', () => {
  testXformHelper(expectations);
});
