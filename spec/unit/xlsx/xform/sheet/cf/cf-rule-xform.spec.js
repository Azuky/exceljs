const testXformHelper = require('../../test-xform-helper');
const XmlStream = verquire('utils/xml-stream');

const CfRuleXform = verquire('xlsx/xform/sheet/cf/cf-rule-xform');

const expectations = [
  {
    title: 'Expression',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'expression',
      dxfId: 1,
      priority: 1,
      formulae: ['ROW()'],
    },
    xml: `
      <cfRule type="expression" dxfId="1" priority="1">
        <formula>ROW()</formula>
      </cfRule>
    `,
    parsedModel: {
      type: 'expression',
      dxfId: 1,
      priority: 1,
      formulae: ['ROW()'],
    },
    tests: ['render', 'parse'],
  },
  {
    title: 'Cell Is',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'cellIs',
      dxfId: 1,
      priority: 1,
      operator: 'greaterThan',
      formulae: ['1'],
    },
    xml: `
      <cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan">
        <formula>1</formula>
      </cfRule>
    `,
    parsedModel: {
      type: 'cellIs',
      dxfId: 1,
      priority: 1,
      operator: 'greaterThan',
      formulae: ['1'],
    },
    tests: ['render', 'parse'],
  },
  {
    title: 'Top 10',
    create() {
      return new CfRuleXform();
    },
    preparedModel: { type: 'top10', dxfId: 1, priority: 1, rank: 10 },
    xml: '<cfRule type="top10" dxfId="1" priority="1" rank="10" />',
    parsedModel: { type: 'top10', dxfId: 1, priority: 1, rank: 10 },
    tests: ['render', 'parse'],
  },
  {
    title: 'Top 10%',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'top10',
      dxfId: 1,
      priority: 1,
      rank: 10,
      percent: true,
    },
    xml: '<cfRule type="top10" dxfId="1" priority="1" rank="10" percent="1" />',
    parsedModel: {
      type: 'top10',
      dxfId: 1,
      priority: 1,
      rank: 10,
      percent: true,
    },
    tests: ['render', 'parse'],
  },
  {
    title: 'Bottom 10',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'top10',
      dxfId: 1,
      priority: 1,
      rank: 10,
      bottom: true,
    },
    xml: '<cfRule type="top10" dxfId="1" priority="1" rank="10" bottom="1" />',
    parsedModel: { type: 'top10', dxfId: 1, priority: 1, rank: 10, bottom: true },
    tests: ['render', 'parse'],
  },
  {
    title: 'Above Average',
    create() {
      return new CfRuleXform();
    },
    preparedModel: { type: 'aboveAverage', dxfId: 1, priority: 1 },
    xml: '<cfRule type="aboveAverage" dxfId="1" priority="1" />',
    parsedModel: { type: 'aboveAverage', dxfId: 1, priority: 1 },
    tests: ['render', 'parse'],
  },
  {
    title: 'Below Average',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'aboveAverage',
      dxfId: 1,
      priority: 1,
      aboveAverage: false,
    },
    xml: '<cfRule type="aboveAverage" dxfId="1" priority="1" aboveAverage="0" />',
    parsedModel: {
      type: 'aboveAverage',
      dxfId: 1,
      priority: 1,
      aboveAverage: false,
    },
    tests: ['render', 'parse'],
  },
  {
    title: 'Colour Scale',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'colorScale',
      priority: 1,
      cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
      color: [{ argb: 'FFFF0000' }, { argb: 'FF00FF00' }, { argb: 'FF0000FF' }],
    },
    xml: `
      <cfRule type="colorScale" priority="1">
        <colorScale>
          <cfvo type="min" />
          <cfvo type="percentile" val="50" />
          <cfvo type="max" />
          <color rgb="FFFF0000" />
          <color rgb="FF00FF00" />
          <color rgb="FF0000FF" />
        </colorScale>
      </cfRule>
    `,
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'parse'],
  },
  {
    title: 'Icon Set',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'iconSet',
      iconSet: '4Arrows',
      priority: 1,
      cfvo: [
        { type: 'percent', value: 0 },
        { type: 'percent', value: 25 },
        { type: 'percent', value: 50 },
        { type: 'percent', value: 75 },
      ],
    },
    xml: `
      <cfRule type="iconSet" priority="1">
        <iconSet iconSet="4Arrows">
          <cfvo type="percent" val="0"/>
          <cfvo type="percent" val="25"/>
          <cfvo type="percent" val="50"/>
          <cfvo type="percent" val="75"/>
        </iconSet>
      </cfRule>
    `,
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'parse'],
  },

  {
    title: 'Data Bar',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'dataBar',
      priority: 1,
      cfvo: [{ type: 'min' }, { type: 'max' }],
      color: { argb: 'FF638EC6' },
      x14Id: '{F7189283-14F7-4DE0-9601-54DE9DB40001}',
    },
    xml: `
      <cfRule type="dataBar" priority="1">
        <dataBar>
          <cfvo type="min"/>
          <cfvo type="max"/>
          <color rgb="FF638EC6"/>
        </dataBar>
        <extLst>
          <ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:id>{F7189283-14F7-4DE0-9601-54DE9DB40001}</x14:id>
          </ext>
        </extLst>
      </cfRule>
    `,
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Data Bar with numeric cfvo',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'dataBar',
      priority: 2,
      cfvo: [
        { type: 'num', value: 5 },
        { type: 'percentile', value: 90.5 },
      ],
      color: { argb: 'FFFF555A' },
      x14Id: '{00000000-0000-0000-0000-000000000002}',
    },
    xml: `
      <cfRule type="dataBar" priority="2">
        <dataBar>
          <cfvo type="num" val="5"/>
          <cfvo type="percentile" val="90.5"/>
          <color rgb="FFFF555A"/>
        </dataBar>
        <extLst>
          <ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:id>{00000000-0000-0000-0000-000000000002}</x14:id>
          </ext>
        </extLst>
      </cfRule>
    `,
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'parse'],
  },
  {
    title: 'Contains Text - formula synthesised from ref and text',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'containsText',
      operator: 'containsText',
      text: 'foo',
      ref: 'B2:D5',
      dxfId: 1,
      priority: 1,
    },
    xml: `
      <cfRule type="containsText" dxfId="1" priority="1">
        <formula>NOT(ISERROR(SEARCH("foo",B2)))</formula>
      </cfRule>
    `,
    parsedModel: {
      type: 'containsText',
      operator: 'containsText',
      dxfId: 1,
      priority: 1,
      formulae: ['NOT(ISERROR(SEARCH("foo",B2)))'],
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Contains Text - explicit formula is used verbatim',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'containsText',
      operator: 'containsText',
      text: 'foo',
      ref: 'B2:D5',
      formulae: ['ISNUMBER(SEARCH("bar",B2))'],
      dxfId: 1,
      priority: 1,
    },
    xml: `
      <cfRule type="containsText" dxfId="1" priority="1">
        <formula>ISNUMBER(SEARCH("bar",B2))</formula>
      </cfRule>
    `,
    parsedModel: {
      type: 'containsText',
      operator: 'containsText',
      dxfId: 1,
      priority: 1,
      formulae: ['ISNUMBER(SEARCH("bar",B2))'],
    },
    tests: ['render', 'parse'],
  },
  ...[
    ['containsBlanks', 'LEN(TRIM(B2))=0'],
    ['notContainsBlanks', 'LEN(TRIM(B2))>0'],
    ['containsErrors', 'ISERROR(B2)'],
    ['notContainsErrors', 'NOT(ISERROR(B2))'],
  ].map(([operator, formula]) => ({
    title: `${operator} - formula synthesised from ref`,
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'containsText',
      operator,
      ref: 'B2:D5',
      dxfId: 1,
      priority: 1,
    },
    // as written by Excel: no `operator` attribute for these cfRule types
    // (ST_ConditionalFormattingOperator has no containsBlanks/containsErrors member)
    xml: `
      <cfRule type="${operator}" dxfId="1" priority="1">
        <formula>${formula}</formula>
      </cfRule>
    `,
    parsedModel: {
      type: 'containsText',
      operator,
      dxfId: 1,
      priority: 1,
      formulae: [formula],
    },
    // known bug: render writes operator="<type>", which is not a valid operator value
    knownFailures: ['render', 'renderIn'],
    knownFailureMessage: /attributes \[dxfId,operator,priority,type\] !== \[dxfId,priority,type\]/,
    tests: ['render', 'renderIn', 'parse'],
  })),
  ...[
    ['today', 'FLOOR(B2,1)=TODAY()'],
    ['yesterday', 'FLOOR(B2,1)=TODAY()-1'],
    ['tomorrow', 'FLOOR(B2,1)=TODAY()+1'],
    ['last7Days', 'AND(TODAY()-FLOOR(B2,1)<=6,FLOOR(B2,1)<=TODAY())'],
    ['thisWeek', 'AND(TODAY()-ROUNDDOWN(B2,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(B2,0)-TODAY()<=7-WEEKDAY(TODAY()))'],
    ['lastWeek', 'AND(TODAY()-ROUNDDOWN(B2,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(B2,0)<(WEEKDAY(TODAY())+7))'],
    ['nextWeek', 'AND(ROUNDDOWN(B2,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(B2,0)-TODAY()<(15-WEEKDAY(TODAY())))'],
    ['thisMonth', 'AND(MONTH(B2)=MONTH(TODAY()),YEAR(B2)=YEAR(TODAY()))'],
    ['lastMonth', 'AND(MONTH(B2)=MONTH(EDATE(TODAY(),0-1)),YEAR(B2)=YEAR(EDATE(TODAY(),0-1)))'],
    ['nextMonth', 'AND(MONTH(B2)=MONTH(EDATE(TODAY(),0+1)),YEAR(B2)=YEAR(EDATE(TODAY(),0+1)))'],
  ].map(([timePeriod, formula]) => ({
    title: `Time Period ${timePeriod} - formula synthesised from ref`,
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'timePeriod',
      timePeriod,
      ref: 'B2:D5',
      dxfId: 1,
      priority: 1,
    },
    xml: `
      <cfRule type="timePeriod" dxfId="1" priority="1" timePeriod="${timePeriod}">
        <formula>${formula.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</formula>
      </cfRule>
    `,
    parsedModel: {
      type: 'timePeriod',
      timePeriod,
      dxfId: 1,
      priority: 1,
      formulae: [formula],
    },
    tests: ['render', 'parse'],
  })),
  {
    title: 'Time Period - explicit formula is used verbatim',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'timePeriod',
      timePeriod: 'today',
      ref: 'B2:D5',
      formulae: ['FLOOR(C3,1)=TODAY()'],
      dxfId: 1,
      priority: 1,
    },
    xml: `
      <cfRule type="timePeriod" dxfId="1" priority="1" timePeriod="today">
        <formula>FLOOR(C3,1)=TODAY()</formula>
      </cfRule>
    `,
    parsedModel: {
      type: 'timePeriod',
      timePeriod: 'today',
      dxfId: 1,
      priority: 1,
      formulae: ['FLOOR(C3,1)=TODAY()'],
    },
    tests: ['render', 'parse'],
  },
  {
    title: 'Time Period - unknown period has no formula',
    create() {
      return new CfRuleXform();
    },
    preparedModel: {
      type: 'timePeriod',
      timePeriod: 'someday',
      ref: 'B2:D5',
      dxfId: 1,
      priority: 1,
    },
    xml: '<cfRule type="timePeriod" dxfId="1" priority="1" timePeriod="someday"/>',
    parsedModel: {
      type: 'timePeriod',
      timePeriod: 'someday',
      dxfId: 1,
      priority: 1,
    },
    tests: ['render', 'parse'],
  },
];

describe('CfRuleXform', () => {
  testXformHelper(expectations);

  describe('synthesised formulas of the blanks and errors rules', () => {
    // the rendered attributes of these rules are a known bug (see above), the formula is not
    const formulaOf = (operator) => {
      const xmlStream = new XmlStream();
      new CfRuleXform().render(xmlStream, { type: 'containsText', operator, ref: 'B2:D5', dxfId: 1, priority: 1 });
      return xmlStream.xml.match(/<formula>(.*)<\/formula>/)[1];
    };

    it('containsBlanks tests the top-left cell for an empty string', () => {
      expect(formulaOf('containsBlanks')).to.equal('LEN(TRIM(B2))=0');
    });
    it('notContainsBlanks tests the top-left cell for a non-empty string', () => {
      expect(formulaOf('notContainsBlanks')).to.equal('LEN(TRIM(B2))&gt;0');
    });
    it('containsErrors tests the top-left cell for an error', () => {
      expect(formulaOf('containsErrors')).to.equal('ISERROR(B2)');
    });
    it('notContainsErrors tests the top-left cell for no error', () => {
      expect(formulaOf('notContainsErrors')).to.equal('NOT(ISERROR(B2))');
    });
  });
});
