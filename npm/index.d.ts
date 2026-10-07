/**
 * TypeScript definitions for excel-community
 */

/** A value that can be written to a cell. Strings starting with `=` are formulas. */
export type CellInput = string | number | boolean | Date | null | undefined;

/** A value read from a cell: dates as `'YYYY-MM-DD'`, date-times as ISO strings, times as `'HH:MM:SS'`. */
export type CellOutput = string | number | boolean | null;

/** Hex color such as `'#1F4E78'`, `'1F4E78'` or `'#FF1F4E78'`. */
export type Color = string;

// ==================== STYLES ====================

export type BorderLineStyle =
  | 'none' | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair'
  | 'dashDot' | 'dashDotDot' | 'mediumDashed' | 'mediumDashDot' | 'mediumDashDotDot' | 'slantDashDot';

export interface BorderStyleOptions {
  style?: BorderLineStyle;
  color?: Color;
}

/** Only the given options change; the rest of the cell style is kept. */
export interface CellStyleOptions {
  bold?: boolean;
  italic?: boolean;
  underline?: 'single' | 'double' | 'none' | boolean;
  strikethrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  fontColor?: Color;
  /** `'none'` or `null` removes the fill. */
  backgroundColor?: Color | 'none' | null;
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'center' | 'bottom';
  wrapText?: boolean;
  shrinkToFit?: boolean;
  rotation?: number;
  /** A built-in format id (0-49) or a custom format code such as `'$#,##0.00'`. */
  numberFormat?: number | string;
  /** Sets the four outer borders. */
  border?: BorderStyleOptions | null;
  leftBorder?: BorderStyleOptions | null;
  rightBorder?: BorderStyleOptions | null;
  topBorder?: BorderStyleOptions | null;
  bottomBorder?: BorderStyleOptions | null;
  diagonalBorder?: BorderStyleOptions | null;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
  /** Cell locking used when the sheet is protected (Excel's default is locked). */
  locked?: boolean | null;
  /** Hides the formula when the sheet is protected. */
  hidden?: boolean | null;
}

export interface BorderInfo {
  style: BorderLineStyle;
  /** `'#RRGGBB'` (`'#AARRGGBB'` when not opaque). */
  color?: string;
}

/** A cell style in the same form `setStyle` takes, so it can be passed back to it. */
export interface CellStyleInfo {
  bold: boolean;
  italic: boolean;
  strikethrough: boolean;
  underline: 'none' | 'single' | 'double';
  fontSize?: number;
  fontFamily?: string;
  /** `'#RRGGBB'` (`'#AARRGGBB'` when not opaque). */
  fontColor: string;
  /** Absent when the cell has no fill. */
  backgroundColor?: string;
  horizontalAlign: 'left' | 'center' | 'right';
  verticalAlign: 'top' | 'center' | 'bottom';
  wrapText: boolean;
  shrinkToFit: boolean;
  rotation: number;
  numberFormat: string;
  leftBorder?: BorderInfo;
  rightBorder?: BorderInfo;
  topBorder?: BorderInfo;
  bottomBorder?: BorderInfo;
  diagonalBorder?: BorderInfo;
  diagonalUp: boolean;
  diagonalDown: boolean;
  locked?: boolean;
  hidden?: boolean;
}

// ==================== HYPERLINKS ====================

/** One target: a web address, an e-mail, a cell of the workbook or a defined name. */
export type HyperlinkTarget = (
  | { url: string }
  | { email: string; subject?: string }
  | { sheet: string; cell?: string }
  | { location: string }
) & { tooltip?: string; display?: string };

export interface HyperlinkOptions {
  /** Text written to the cell. */
  text?: string;
  /** Apply Excel's hyperlink look (blue, underlined). Default `true`. */
  styled?: boolean;
}

export interface HyperlinkInfo {
  url?: string;
  location?: string;
  tooltip?: string;
  display?: string;
}

// ==================== CHARTS ====================

export interface ChartSeriesStyleOptions {
  fillColor?: Color;
  fillType?: 'solid' | 'transparent' | 'none';
  /** Opacity of a `transparent` fill, 0-100. Default 50. */
  fillAlpha?: number;
  borderColor?: Color;
  /** Opacity of the border, 0-100. Default 100. */
  borderAlpha?: number;
  /** Line width in EMUs, e.g. `'28575'` (2.25 pt). */
  borderWidth?: string | number;
}

export interface ChartSeriesConfig {
  name?: string;
  categoriesRange: string;
  valuesRange: string;
  /** Bubble charts: the range with the bubble sizes. */
  bubbleSizeRange?: string;
  /** Shortcut for `style.fillColor`. */
  colorHex?: Color;
  style?: ChartSeriesStyleOptions;
}

export interface ChartDataLabelsOptions {
  value?: boolean;
  categoryName?: boolean;
  seriesName?: boolean;
  /** Pie and doughnut charts only. */
  percentage?: boolean;
  separator?: string;
  /** `'outEnd'`, `'ctr'`, `'inEnd'`, `'t'`, `'b'`, `'bestFit'`, ... */
  labelPosition?: string;
}

/** Either `{column, row, width, height}` (in cells) or `{fromCol, fromRow, toCol, toRow}`. */
export type ChartAnchorOptions =
  | { column?: number; row?: number; width?: number; height?: number }
  | { fromCol?: number; fromRow?: number; toCol?: number; toRow?: number };

export type ChartType =
  | 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'ofPie' | 'pieOfPie' | 'barOfPie'
  | 'scatter' | 'bubble' | 'stock' | 'radar';

export interface ChartConfig {
  type?: ChartType;
  title?: string;
  showLegend?: boolean;
  series: ChartSeriesConfig[];
  anchor?: ChartAnchorOptions;
  dataLabels?: ChartDataLabelsOptions;
  /** Column, bar, line and area charts. */
  grouping?: 'clustered' | 'stacked' | 'percentStacked';
  /** Line and scatter charts. */
  showMarkers?: boolean;
  smooth?: boolean;
  /** Scatter charts. */
  showLines?: boolean;
  /** Radar charts. */
  filled?: boolean;
  /** Bubble charts. */
  bubbleScale?: number;
  showNegativeBubbles?: boolean;
  /** Stock charts (series: open?, high, low, close). */
  showHighLowLines?: boolean;
  showUpDownBars?: boolean;
  /** Of-pie charts. */
  ofPieType?: 'pie' | 'bar';
  splitType?: 'position' | 'value' | 'percent' | 'custom';
  splitPosition?: number;
  secondPieSize?: number;
}

// ==================== CONDITIONAL FORMATTING ====================

export interface ConditionalStyle {
  backgroundColor?: Color;
  fontColor?: Color;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  underline?: 'single' | 'double' | 'none' | boolean;
}

export type ConditionalOperator =
  | 'equal' | 'notEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual'
  | 'between' | 'notBetween';

export interface ConditionalRule {
  type:
    | 'cellIs' | 'expression' | 'containsText' | 'notContains' | 'beginsWith' | 'endsWith'
    | 'duplicateValues' | 'uniqueValues';
  /** `cellIs` rules. */
  operator?: ConditionalOperator;
  /** `cellIs` rules: a number or formula; `value2` is the upper bound of `between`. */
  value?: number | string;
  value2?: number | string;
  formulae?: (number | string)[];
  /** `expression` rules, e.g. `'$C2>100'`. */
  formula?: string;
  /** Text rules. */
  text?: string;
  style?: ConditionalStyle;
  priority?: number;
}

export interface ConditionalFormattingInfo {
  range: string;
  rules: (Omit<ConditionalRule, 'value' | 'value2'> & { formulae: string[] })[];
}

// ==================== DATA VALIDATION ====================

export type ValidationOperator =
  | 'between' | 'notBetween' | 'equal' | 'notEqual' | 'lessThan' | 'lessThanOrEqual'
  | 'greaterThan' | 'greaterThanOrEqual';

interface ValidationMessages {
  prompt?: { title?: string; message?: string };
  error?: { title?: string; message?: string; style?: 'stop' | 'warning' | 'information' };
  allowBlank?: boolean;
  showDropdown?: boolean;
}

export type DataValidationRule = ValidationMessages &
  (
    | { type: 'list'; items: string[] }
    | { type: 'listFromRange'; range: string; sheet?: string }
    | { type: 'wholeNumber' | 'decimal' | 'textLength'; operator?: ValidationOperator; value: number; value2?: number }
    | { type: 'date'; operator?: ValidationOperator; value: Date | string; value2?: Date | string }
    /** Times as `'HH:MM[:SS]'`. */
    | { type: 'time'; operator?: ValidationOperator; value: string; value2?: string }
    /** Formula without `=` relative to the first cell, e.g. `'H2>E2'`. */
    | { type: 'custom'; formula: string }
    | { type: 'inputMessage' }
  );

export interface DataValidationInfo {
  type: string;
  operator: string;
  formula1?: string;
  formula2?: string;
  items?: string[];
  allowBlank: boolean;
  showDropdown: boolean;
  prompt?: { title?: string; message: string };
  error?: { title?: string; message: string; style: string };
}

// ==================== PAGE SETUP ====================

export type PaperSizeName =
  | 'letter' | 'tabloid' | 'ledger' | 'legal' | 'statement' | 'executive' | 'a2' | 'a3' | 'a4' | 'a5' | 'a6'
  | 'b4' | 'b5' | 'folio' | 'quarto' | 'envelope10' | 'envelopeDL' | 'envelopeC5';

/** Only the given options change. */
export interface PageSetupOptions {
  orientation?: 'portrait' | 'landscape' | 'automatic';
  /** A paper name or Excel's numeric paper code. */
  paperSize?: PaperSizeName | number;
  /** Print scale in percent (10-400). */
  scale?: number;
  /** Fit to N pages wide / tall (0 = as many as needed). */
  fitToWidth?: number;
  fitToHeight?: number;
  firstPageNumber?: number;
  pageOrder?: 'downThenOver' | 'overThenDown';
  blackAndWhite?: boolean;
  draft?: boolean;
  cellComments?: 'none' | 'atEnd' | 'asDisplayed';
  errors?: 'displayed' | 'blank' | 'dash' | 'na';
  copies?: number;
}

export interface PageSetupInfo extends Omit<PageSetupOptions, 'paperSize'> {
  paperSize?: string;
  paperSizeCode?: number;
  fitToPage?: boolean;
}

export type PageMarginsOptions =
  | 'normal'
  | 'wide'
  | 'narrow'
  | { left?: number; right?: number; top?: number; bottom?: number; header?: number; footer?: number; unit?: 'in' | 'cm' };

export interface PageMarginsInfo {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

export interface PrintOptions {
  gridLines?: boolean;
  headings?: boolean;
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
}

/** Text with Excel codes: `&L`/`&C`/`&R` sections, `&P` page, `&N` pages, `&D` date, `&A` sheet. */
export interface HeaderFooterOptions {
  header?: string;
  footer?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
  differentFirst?: boolean;
  differentOddEven?: boolean;
  scaleWithDoc?: boolean;
  alignWithMargins?: boolean;
}

export interface HeaderFooterInfo {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
  differentFirst?: boolean;
  differentOddEven?: boolean;
}

// ==================== TABLES ====================

export type TableTotalsFunction =
  | 'none' | 'sum' | 'average' | 'count' | 'countNumbers' | 'min' | 'max' | 'stdDev' | 'variance' | 'custom';

export interface TableColumnOptions {
  name: string;
  totalsFunction?: TableTotalsFunction;
  totalsLabel?: string;
  /** Used with `totalsFunction: 'custom'`. */
  totalsRowFormula?: string;
  calculatedColumnFormula?: string;
}

export interface TableOptions {
  /** `'TableStyleMedium9'`, `'medium9'`, `'light1'`, `'dark2'`, or `null`/`'none'`. Default `'medium2'`. */
  style?: string | null;
  showHeaderRow?: boolean;
  showTotalsRow?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showFilterButtons?: boolean;
}

export interface TableUpdateOptions extends TableOptions {
  name?: string;
  ref?: string;
  columns?: (string | TableColumnOptions)[];
}

export interface TableInfo extends Required<Omit<TableOptions, 'style'>> {
  name: string;
  ref: string;
  style: string | null;
  columns: (Required<Pick<TableColumnOptions, 'name' | 'totalsFunction'>> & Omit<TableColumnOptions, 'name' | 'totalsFunction'>)[];
  dataRowCount: number;
}

// ==================== PIVOT TABLES ====================

export type PivotFunction =
  | 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'stdDevp' | 'var' | 'varp';

export interface PivotTableConfig {
  name?: string;
  sourceSheet: string;
  /** Source data including the header row, e.g. `'A1:C100'`. */
  sourceRange: string;
  /** Top-left cell of the pivot table on its sheet. */
  targetCell?: string;
  /** Header names used as row / column fields. */
  rows?: string[];
  columns?: string[];
  values: (string | { field: string; function?: PivotFunction; customName?: string })[];
}

// ==================== OTHER OPTIONS ====================

/** `true` allows the action while the sheet is protected; options left out are blocked. */
export interface SheetProtectionOptions {
  objects?: boolean;
  scenarios?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
}

export interface FilterColumnOptions {
  /** 0-based column offset inside the AutoFilter range. */
  column: number;
  values?: string[];
  blank?: boolean;
  custom?: {
    operator: 'equal' | 'notEqual' | 'lessThan' | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual';
    value: string | number;
  }[];
  /** Combine `custom` rules with AND instead of OR. */
  and?: boolean;
}

export interface AutoFilterInfo {
  ref: string;
  columns: { column: number; values: string[]; blank: boolean; custom: { operator: string; value: string }[]; and: boolean }[];
}

export interface OutlineGroup {
  start: number;
  end: number;
  level: number;
  collapsed: boolean;
}

export interface OutlineSettings {
  summaryBelow?: boolean;
  summaryRight?: boolean;
  showOutlineSymbols?: boolean;
  applyStyles?: boolean;
}

export interface ExportOptions {
  /** 0-based row holding the column names. Default 0. */
  headerRow?: number;
  /** `'typed'` (numbers, booleans, `Date`s) or `'displayText'` (as Excel shows them). */
  mode?: 'typed' | 'displayText';
  skipEmptyRows?: boolean;
}

export interface CsvOptions {
  separator?: string;
  lineTerminator?: string;
  /** Default `'displayText'`. */
  mode?: 'typed' | 'displayText';
  skipEmptyRows?: boolean;
}

export type ExportValue = string | number | boolean | Date | null;

export interface FindReplaceOptions {
  /** Stop after this many replacements. */
  first?: number;
  startingRow?: number;
  endingRow?: number;
  startingColumn?: number;
  endingColumn?: number;
}

// ==================== CELL / SHEET / WORKBOOK ====================

export interface Cell {
  readonly cellId: string;
  readonly row: number;
  readonly col: number;
  readonly displayText: string;
  readonly type: 'string' | 'int' | 'double' | 'bool' | 'date' | 'datetime' | 'time' | 'formula' | 'null';
  readonly style: CellStyleInfo | null;
  get value(): CellOutput;
  set value(value: CellInput);
  /** A JS `Date` for date and date-time cells. */
  readonly dateValue: Date | null;
  /** The result Excel cached for a formula cell in the file. */
  readonly cachedValue: CellOutput;
  formula: string | null;
  comment: string | null;
  setFormula(formula: string): void;
  /** Stores a time of day: `setTime('14:30')` or `setTime(14, 30, 0)`. */
  setTime(time: string): void;
  setTime(hour: number, minute?: number, second?: number): void;
  /** Changes only the given options; the rest of the style is kept. */
  setStyle(options: CellStyleOptions): void;
  resetStyle(): void;
  setHyperlink(url: string, tooltip?: string, text?: string): void;
  setHyperlink(target: HyperlinkTarget, options?: HyperlinkOptions): void;
  getHyperlink(): HyperlinkInfo | null;
  removeHyperlink(): void;
  readonly dataValidation: DataValidationInfo | null;
  /** Whether a value passes the cell's validation (`null` when it depends on formulas). */
  validates(value: CellInput): boolean | null;
}

export interface Sheet {
  readonly name: string;
  readonly maxRows: number;
  readonly maxColumns: number;
  rightToLeft: boolean;
  /** `'#RRGGBB'`; `null` when there is none or it is a theme color. */
  tabColor: string | null;
  frozenRows: number | null;
  frozenColumns: number | null;
  readonly spannedItems: string[];
  readonly rows: (Cell | null)[][];

  // Cells, rows and columns
  cell(cellIndex: string): Cell;
  appendRow(values: CellInput[]): void;
  /** Appends several rows in one call; faster than calling `appendRow` for each. */
  appendRows(rows: CellInput[][]): void;
  insertRowIterables(values: CellInput[], rowIndex: number, options?: { startingColumn?: number; overwriteMergedCells?: boolean }): void;
  insertRow(rowIndex: number): void;
  removeRow(rowIndex: number): void;
  insertColumn(columnIndex: number): void;
  removeColumn(columnIndex: number): void;
  clearRow(rowIndex: number): boolean;
  /** Values of a range such as `'A1:C3'`. */
  rangeValues(range: string): CellOutput[][];
  /** Returns the number of replacements. */
  findAndReplace(source: string | RegExp, target: string, options?: FindReplaceOptions): number;

  // Dimensions & visibility
  setColumnHidden(col: number, hidden: boolean): void;
  isColumnHidden(col: number): boolean;
  setRowHidden(row: number, hidden: boolean): void;
  isRowHidden(row: number): boolean;
  setColumnWidth(col: number, width: number): void;
  getColumnWidth(col: number): number;
  setRowHeight(row: number, height: number): void;
  getRowHeight(row: number): number;
  setDefaultColumnWidth(width: number): void;
  setDefaultRowHeight(height: number): void;
  setColumnAutoFit(col: number): void;
  /** Office theme color index (0-11) with an optional tint (-1 to 1). */
  setTabColorTheme(theme: number, tint?: number): void;

  // Merging
  merge(startCell: string, endCell: string, customValue?: CellInput): void;
  /** Unmerges a range (`'A1:E1'`) or the merged range containing a cell (`'A1'`). */
  unmerge(cellRef: string): void;

  // AutoFilter
  readonly hasAutoFilter: boolean;
  readonly autoFilter: AutoFilterInfo | null;
  setAutoFilter(range: string): void;
  addFilterColumn(options: FilterColumnOptions): void;
  clearAutoFilter(): void;

  // Protection
  readonly isProtected: boolean;
  protect(password?: string, options?: SheetProtectionOptions): void;
  unprotect(): void;

  // Grouping
  groupRows(start: number, end: number, options?: { collapsed?: boolean }): void;
  ungroupRows(start: number, end: number): void;
  groupColumns(start: number, end: number, options?: { collapsed?: boolean }): void;
  ungroupColumns(start: number, end: number): void;
  collapseRowGroup(start: number, end: number): void;
  expandRowGroup(start: number, end: number): void;
  collapseColumnGroup(start: number, end: number): void;
  expandColumnGroup(start: number, end: number): void;
  clearGrouping(): void;
  getRowOutlineLevel(row: number): number;
  getColumnOutlineLevel(col: number): number;
  readonly rowGroups: OutlineGroup[];
  readonly columnGroups: OutlineGroup[];
  /** Assigning changes only the given settings. */
  get outlineSettings(): Required<OutlineSettings>;
  set outlineSettings(settings: OutlineSettings);

  // Images & charts
  addImage(
    bytes: Uint8Array,
    format: string,
    col: number,
    row: number,
    widthPixels: number,
    heightPixels: number,
    options?: { colOffset?: number; rowOffset?: number }
  ): void;
  addChart(config: ChartConfig | string): void;
  readonly chartCount: number;

  // Conditional formatting
  addConditionalFormatting(range: string, rules: ConditionalRule | ConditionalRule[]): void;
  clearConditionalFormatting(): void;
  readonly conditionalFormattings: ConditionalFormattingInfo[];

  // Data validation
  addDataValidation(range: string, rule: DataValidationRule): void;
  removeDataValidation(range: string): void;
  clearDataValidations(): void;
  getDataValidation(cellRef: string): DataValidationInfo | null;
  readonly dataValidations: Record<string, DataValidationInfo>;

  // Hyperlinks
  setHyperlinkRange(range: string, target: HyperlinkTarget, options?: { styled?: boolean }): void;
  clearHyperlinks(): void;
  readonly hyperlinks: Record<string, HyperlinkInfo>;

  // Page setup & printing
  setPageSetup(options: PageSetupOptions): void;
  readonly pageSetup: PageSetupInfo | null;
  clearPageSetup(): void;
  setPageMargins(margins: PageMarginsOptions): void;
  readonly pageMargins: PageMarginsInfo | null;
  clearPageMargins(): void;
  setPrintOptions(options: PrintOptions): void;
  readonly printOptions: PrintOptions | null;
  clearPrintOptions(): void;
  setHeaderFooter(options: HeaderFooterOptions): void;
  readonly headerFooter: HeaderFooterInfo | null;
  clearHeaderFooter(): void;

  // Tables
  addTable(range: string, name: string, columns?: (string | TableColumnOptions)[] | string | null, options?: TableOptions): TableInfo;
  getTable(name: string): TableInfo | null;
  readonly tables: TableInfo[];
  updateTable(name: string, options: TableUpdateOptions): void;
  removeTable(name: string): void;
  tableRowsAsMaps(name: string, options?: { mode?: 'typed' | 'displayText' }): Record<string, ExportValue>[];
  appendTableRow(name: string, values: CellInput[]): TableInfo;

  // Pivot tables
  addPivotTable(config: PivotTableConfig): void;
  readonly pivotTableCount: number;
  /** Recomputes the pivot table cells after their source data changed (saving does it too). */
  refreshPivotTables(): void;

  // Export & import
  rowsAsMaps(options?: ExportOptions): Record<string, ExportValue>[];
  rowsAsValues(options?: Omit<ExportOptions, 'headerRow'>): ExportValue[][];
  toJson(options?: ExportOptions & { indent?: string }): string;
  toCsv(options?: CsvOptions): string;
  appendRowsFromMaps(rows: Record<string, CellInput>[], options?: { headerRow?: number; writeHeader?: boolean }): void;
}

export interface Workbook {
  readonly sheets: string[];
  defaultSheet: string | null;
  sheet(name: string): Sheet;
  createSheet(name: string): Sheet;
  deleteSheet(name: string): boolean;
  renameSheet(oldName: string, newName: string): boolean;
  copySheet(fromSheet: string, toSheet: string): boolean;
  /** Makes `name` share its content with `existingSheet`. */
  linkSheet(name: string, existingSheet: string): void;
  unlinkSheet(name: string): void;
  /** Every sheet as `{sheetName: rows}`. */
  toMaps(options?: ExportOptions): Record<string, Record<string, ExportValue>[]>;
  toJson(options?: ExportOptions & { indent?: string }): string;
  encode(): Uint8Array | null;
  encodeBase64(): string | null;
  /** A Node.js `Buffer` when available, otherwise a `Uint8Array`. */
  toBuffer(): Uint8Array | null;
  /** Writes the file in Node.js or downloads it in a browser. */
  save(fileName?: string): string | Uint8Array;
}

// ==================== ERRORS ====================

/** Base class of the errors thrown by the library; the original error is in `cause`. */
export class ExcelError extends Error {
  readonly name: 'ExcelError' | 'ExcelArgumentError' | 'ExcelStateError' | 'ExcelFormatError';
}
/** An invalid argument or option: bad cell reference, range, value... */
export class ExcelArgumentError extends ExcelError {
  readonly name: 'ExcelArgumentError';
}
/** The operation is not possible in the current state. */
export class ExcelStateError extends ExcelError {
  readonly name: 'ExcelStateError';
}
/** The data given to `Excel.read` is not a readable workbook. */
export class ExcelFormatError extends ExcelError {
  readonly name: 'ExcelFormatError';
}

export class Excel {
  static create(): Workbook;
  /** Reads `.xlsx` or legacy `.xls` bytes, or a Base64 string. */
  static read(data: Uint8Array | ArrayBuffer | ArrayBufferView | string): Workbook;
  static fromFile(filePath: string): Workbook;
}

export default Excel;
