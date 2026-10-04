/**
 * TypeScript definitions for excel-community
 */

export interface BorderStyleOptions {
  style?: 'none' | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair' | 'dashdot' | 'dashdotdot';
  color?: string;
}

export interface CellStyleOptions {
  bold?: boolean;
  italic?: boolean;
  underline?: 'single' | 'double' | boolean;
  strikethrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  fontColor?: string;
  backgroundColor?: string;
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'center' | 'bottom';
  wrapText?: boolean;
  rotation?: number;
  numberFormat?: number | string;
  border?: BorderStyleOptions;
  leftBorder?: BorderStyleOptions;
  rightBorder?: BorderStyleOptions;
  topBorder?: BorderStyleOptions;
  bottomBorder?: BorderStyleOptions;
}

export interface HyperlinkInfo {
  url?: string;
  location?: string;
  tooltip?: string;
  display?: string;
}

export interface SheetProtectionOptions {
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  autoFilter?: boolean;
  sort?: boolean;
}

export interface Cell {
  readonly cellId: string;
  readonly row: number;
  readonly col: number;
  readonly displayText: string;
  readonly type: string;
  readonly style: any;
  value: any;
  formula: string | null;
  comment: string | null;
  setFormula(formula: string): void;
  setStyle(options: CellStyleOptions): void;
  setHyperlink(url: string, tooltip?: string, display?: string): void;
  getHyperlink(): HyperlinkInfo | null;
  removeHyperlink(): void;
}

export interface Sheet {
  readonly name: string;
  readonly maxRows: number;
  readonly maxColumns: number;
  rightToLeft: boolean;
  tabColor: string | null;
  frozenRows: number | null;
  frozenColumns: number | null;
  readonly spannedItems: string[];
  readonly hasAutoFilter: boolean;
  readonly rows: (Cell | null)[][];

  setColumnHidden(col: number, hidden: boolean): void;
  isColumnHidden(col: number): boolean;
  setRowHidden(row: number, hidden: boolean): void;
  isRowHidden(row: number): boolean;

  cell(cellIndex: string): Cell;
  appendRow(values: (string | number | boolean | null)[]): void;
  setColumnWidth(col: number, width: number): void;
  getColumnWidth(col: number): number;
  setRowHeight(row: number, height: number): void;
  getRowHeight(row: number): number;
  setDefaultColumnWidth(width: number): void;
  setDefaultRowHeight(height: number): void;
  setColumnAutoFit(col: number): void;
  merge(startCell: string, endCell: string): void;
  unmerge(cellRef: string): void;
  setAutoFilter(range: string): void;
  clearAutoFilter(): void;
  protect(password?: string, options?: SheetProtectionOptions): void;
  unprotect(): void;
  groupRows(start: number, end: number): void;
  ungroupRows(start: number, end: number): void;
  groupColumns(start: number, end: number): void;
  ungroupColumns(start: number, end: number): void;
  addImage(bytes: Uint8Array, format: string, col: number, row: number, widthPixels: number, heightPixels: number): void;
  addChart(chartConfigJson: string): void;
  addTable(range: string, name: string, columnsJson?: string): void;
}

export interface Workbook {
  readonly sheets: string[];
  defaultSheet: string | null;
  sheet(name: string): Sheet;
  createSheet(name: string): Sheet;
  deleteSheet(name: string): boolean;
  renameSheet(oldName: string, newName: string): boolean;
  copySheet(fromSheet: string, toSheet: string): boolean;
  encode(): Uint8Array | null;
  encodeBase64(): string | null;
  toBuffer(): Buffer | Uint8Array;
  save(fileName?: string): string | Uint8Array;
}

export class Excel {
  static create(): Workbook;
  static read(data: Uint8Array | Buffer | string): Workbook;
  static fromFile(filePath: string): Workbook;
}

export default Excel;
