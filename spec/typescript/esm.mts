import ExcelJS, { Workbook, ValueType } from 'exceljs';

const wb: InstanceType<typeof Workbook> = new ExcelJS.Workbook();
const v: typeof ValueType = ExcelJS.ValueType;
void wb;
void v;
