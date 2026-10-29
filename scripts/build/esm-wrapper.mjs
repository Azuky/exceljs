// ES module wrapper: the only lib code lives in the CJS bundle (dist/index.cjs),
// so require() and import share the same classes (no dual-package hazard).
// './index.cjs' is resolved relative to dist/ (see tsdown.config.ts).
import ExcelJS from './index.cjs';

export const {
  Workbook,
  ModelContainer,
  stream,
  ValueType,
  FormulaType,
  RelationshipType,
  DocumentType,
  ReadingOrder,
  ErrorValue,
} = ExcelJS;

export default ExcelJS;
