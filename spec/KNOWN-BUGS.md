# Known lib bugs and test conventions

## Test conventions

- **Known bugs**: `knownBug(title, fn, /reason/)` (`spec/utils/known-bug.js`) registers a test whose body states the
  CORRECT behaviour. It passes while the body throws an error whose message matches `reason`, and goes red when the bug
  is fixed (no error) or breaks differently. When it goes red after a fix, remove the marker and keep the body as a
  normal `it`.
- **Known deviations (xforms)**: an expectation in `spec/unit/xlsx/xform/**` can declare `knownDeviations` (see
  `spec/unit/xlsx/xform/test-xform-helper.js`) per test kind (`render`, `parse`; `renderIn`/`parseIn` follow them).
  `neutralise` undoes exactly the deviation so the rest of the output is still compared strictly; a second test
  `"... (known deviation)"` checks that the raw comparison still fails with `reason`. It goes red when the lib is fixed:
  drop the declaration then.
- `reason` regexes pin chai's message text, so a chai or vitest bump may require updating them.
- **Assertion style**: chai (`expect(...).to...`) everywhere. The one allowed exception is vitest spy matchers
  (`toHaveBeenCalled*`) together with `vi.spyOn`, because chai has no spy matchers.
- **XML comparison**: `spec/utils/chai-xml-ordered.js` compares sibling element order (the old chai-xml did not).
  Expected XML follows the ECMA-376 schema order.

## Known lib bugs pinned by tests

Format: lib location - wrong vs expected - spec file: test title. `knownBug` tests are listed by their title,
`knownDeviations` entries by the expectation title.

### Utils, doc and CSV

- `lib/utils/stream-buf.js:123` - `read(size)` copies up to index `size` instead of `iRead + size`; a second partial
  read returns NUL bytes - `spec/unit/utils/stream-buf.spec.js`: reads a buffer back in consecutive slices.
- `lib/doc/cell.js:827-850` - `FormulaValue.effectiveType` has no branch for boolean or error results and reports
  Null - `spec/unit/doc/cell.spec.js`: is Boolean for a boolean result; is Error for an error result.
- `lib/doc/cell.js:531` - `RichTextValue.toCsvString` reads `this.text`, which does not exist, and throws -
  `spec/unit/doc/cell.spec.js`: toCsvString renders a rich text cell as its quoted plain text.
- `lib/doc/table.js:407-411` - setter is named `displayNamename` and getter reads `displyName`, so assigning
  `displayName` has no effect - `spec/unit/doc/worksheet-table.spec.js`: stores a new display name.
- `lib/doc/anchor.js:37` - `asInstance` calls `new Anchor(model)`, the model lands in the worksheet parameter and is
  ignored - `spec/unit/doc/anchor.spec.js`: asInstance wraps a plain model into an Anchor with the same position.
- `lib/doc/range.js:22` - `setTLBR(tl, br, sheetName)` re-applies the corners with an undefined sheet name, dropping
  it - `spec/unit/doc/range.spec.js`: takes the sheet name from the third argument of [tl, br, sheetName].
- `lib/stream/xlsx/workbook-writer.js:174-176` - `getWorksheet()` uses `find(() => true)` on a sparse array whose
  index 0 is a hole, returning undefined - `spec/unit/doc/workbook-writer.spec.js`: returns the first sheet when called
  without arguments.
- `lib/stream/xlsx/workbook-writer.js:136` - `addWorksheet` does no name validation: illegal characters and
  case-insensitive duplicates are accepted - `spec/unit/doc/workbook-writer.spec.js`: rejects names with illegal
  characters; rejects a duplicate name, ignoring case.
- `lib/csv/csv.js:134` - `value.result || ''` writes falsy formula results (0, false) as empty fields -
  `spec/integration/workbook/workbook.spec.js`: writes falsy formula results.
- `lib/csv/csv.js:146` - rich text cells are written as a JSON blob instead of their text (may be by design, see
  open questions) - `spec/integration/workbook/workbook.spec.js`: writes a rich text cell as its plain concatenated
  text.
- `lib/csv/csv.js:49-69` - the default `YYYY-MM-DD[T]HH:mm:ssZ` format is matched strictly against the local offset,
  so a UTC `...Z` timestamp (what `write` emits with `dateUTC`) stays a string -
  `spec/integration/workbook/workbook.spec.js`: reads back the UTC timestamps that write emits with dateUTC.
- `lib/csv/csv.js:88` - `parserOptions.headers` makes fast-csv emit objects but `read()` calls `data.map` on each row,
  so every read rejects - `spec/integration/workbook/workbook.spec.js`: reads with parserOptions.headers.

### XLSX xforms

- `lib/xlsx/xform/table/table-column-xform.js:18,30` - column dxf is written/read as `dxfId`; the schema attribute
  is `dataDxfId` - `spec/unit/xlsx/xform/table/table-xform.spec.js`: parses the dataDxfId of a table column; and
  deviations (render + parse) of column with dxf style.
- `lib/xlsx/xform/sheet/cf/cf-rule-xform.js:75-88,239-243` - the `text` attribute of a containsText rule is neither
  read nor written, and `operator="containsText"` is omitted on write -
  `spec/unit/xlsx/xform/sheet/cf/cf-rule-xform.spec.js`: parses the text attribute of a containsText rule; and
  deviations (render) of Contains Text - formula synthesised from ref and text; Contains Text - explicit formula is
  used verbatim.
- `lib/xlsx/xform/sheet/cf/cf-rule-xform.js:239-243` - `renderText` writes `operator="<type>"` for containsBlanks,
  notContainsBlanks, containsErrors and notContainsErrors, which is not a valid operator value; deviation (render) in
  `spec/unit/xlsx/xform/sheet/cf/cf-rule-xform.spec.js`: `${operator} - formula synthesised from ref`.
- `lib/xlsx/xform/sheet/cf-ext/icon-set-ext-xform.js:35` - render overwrites `icon.iconId` with the list position -
  deviation (render) in `spec/unit/xlsx/xform/sheet/cf-ext/icon-set-ext-xform.spec.js`: Custom icons keep their own
  iconId.
- `lib/xlsx/xform/drawing/one-cell-anchor-xform.js:29` - writes `editAs` on `xdr:oneCellAnchor`
  (CT_OneCellAnchor has no attributes) - deviations (render) in
  `spec/unit/xlsx/xform/drawing/one-cell-anchor-xform.spec.js`: Picture anchored at one cell with a size;
  `spec/unit/xlsx/xform/drawing/drawing-xform.spec.js`: Drawing 1.
- `lib/xlsx/xform/sheet/sheet-properties-xform.js:28-29` - renders `pageSetUpPr` before `outlinePr`; CT_SheetPr
  requires tabColor, outlinePr, pageSetUpPr - deviations (render) in
  `spec/unit/xlsx/xform/sheet/sheet-properties-xform.spec.js`: pageSetup + outlineProperties; tabColor +
  outlineProperties + pageSetup.
- `lib/xlsx/xform/sheet/worksheet-xform.js:352-360` - `legacyDrawing` is written after `tableParts` (and extLst);
  schema order is the reverse - deviation (render) in `spec/unit/xlsx/xform/sheet/worksheet-xform.spec.js`: Sheet 8 -
  Comments and Table (legacyDrawing).
- `lib/xlsx/xform/sheet/page-breaks-xform.js:14` - `parseOpen` reads the non-existent `ref` attribute of `<brk>`
  instead of id/min/max/man; deviation (parse) in `spec/unit/xlsx/xform/sheet/page-breaks-xform.spec.js`: one page
  break.
- `lib/xlsx/xform/sheet/row-breaks-xform.js` (via `page-breaks-xform.js:14`) - the `<brk>` children yield no model, so
  the list stays empty; deviation (parse) in `spec/unit/xlsx/xform/sheet/row-breaks-xform.spec.js`: two manual row
  breaks.

### XLSX container and streaming

- `lib/utils/zip-stream.js:26-27`, `lib/stream/xlsx/workbook-writer.js:243` - the `base64` entry option is not honoured
  by the zip layer, so the base64 text is stored instead of the image bytes -
  `spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js`: with background image from base64 stores the
  decoded bytes.
- `lib/stream/xlsx/workbook-reader.js:50-53,271-272` - `_parseSharedStrings` yields bare `{ index, text }` objects but
  `read()` destructures `{ eventType, value }`, so nothing is emitted -
  `spec/integration/workbook-xlsx-stream-reader.spec.js`: sharedStrings:emit emits a 'shared-strings' event per string;
  parse() yields 'shared-strings' events.
- `lib/stream/xlsx/workbook-reader.js:205-212` - run colour: reads `attributes.argb` instead of `rgb`, and the theme
  stays a string - same spec: reads the rgb colour of a run as argb; reads the theme colour of a run as a number.
- `lib/stream/xlsx/workbook-reader.js:229` - run font name reads `node.value` instead of `node.attributes.val` - same
  spec: reads the font name of a run.
- `lib/stream/xlsx/workbook-reader.js:198-200` - run charset reads `attributes.charset` instead of `attributes.val` -
  same spec: reads the charset of a run.
- `lib/stream/xlsx/workbook-reader.js:240-241` - the `strike` case is an empty break - same spec: reads strike of a
  run.
- `lib/stream/xlsx/hyperlink-reader.js:17-22` - the cache is an object keyed by rId but `count` reads `.length` and
  `each` calls `.forEach` - same spec: count is the number of cached hyperlinks; each iterates the cached hyperlinks.
- `lib/stream/xlsx/worksheet-reader.js:296,318` and `lib/xlsx/xform/sheet/cell-xform.js:351,369` - text already
  decoded by the XML parser is decoded again with `utils.xmlDecode` (strings and string formula results), so
  `&amp;` becomes `&` - `spec/integration/workbook-xlsx-stream-reader.spec.js`: does not decode entities twice; does
  not decode entities twice in a string formula result; `spec/integration/workbook/workbook.spec.js`: keeps a string
  containing an entity reference; keeps a string formula result containing an entity reference.
- `lib/stream/xlsx/workbook-reader.js:43,141` - `fs.createReadStream` without an error listener: a missing file is an
  uncaught exception, and with a listener the reader never reports anything -
  `spec/integration/workbook-xlsx-stream-reader.spec.js`: emits 'error' for a file that does not exist.

## Open questions for the owner

- CSV rich text is written as JSON (`lib/csv/csv.js:146`, UD-4): bug or by design?
- `StreamBuf` accepts writes after `end()` (characterisation test in `spec/unit/utils/stream-buf.spec.js`, "write after
  end"): intended?
- Stream reader defaults: README documents `entries: 'emit'` and `hyperlinks: 'cache'`, but the code defaults both to
  `'ignore'` (`lib/stream/xlsx/workbook-reader.js:28,30`). Which one should win?
