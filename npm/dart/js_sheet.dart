import 'dart:js_interop';
import 'dart:js_interop_unsafe';
import 'package:excel_community/excel_community.dart';
import 'js_cell.dart';
import 'js_convert.dart';
import 'js_options.dart';

@JSExport()
class JsSheet {
  final Sheet _sheet;
  JsSheet(this._sheet);

  String get name => _sheet.sheetName;
  int get maxRows => _sheet.maxRows;
  int get maxColumns => _sheet.maxColumns;

  bool get rightToLeft => _sheet.isRTL;
  set rightToLeft(bool val) => _sheet.isRTL = val;

  // ==================== TAB COLOR ====================

  String? get tabColor => colorInfo(_sheet.tabColor?.rgb ?? _sheet.tabColor?.color?.colorHex);
  set tabColor(String? hex) {
    if (hex == null || hex.isEmpty) {
      _sheet.clearTabColor();
    } else {
      _sheet.setTabColorHex(hex);
    }
  }

  /// Office theme color index (0-11) with an optional tint (-1.0 to 1.0).
  void setTabColorTheme(int theme, [double? tint]) {
    _sheet.tabColor = TabColor.fromTheme(theme, tint: tint);
  }

  // ==================== CELLS & ROWS ====================

  /// Creates the cell if needed and returns its position as
  /// `row * 16384 + column`; wrapper.js turns it into a `Cell`.
  int cell(String cellIndex) {
    final data = _sheet.cell(CellIndex.indexByString(cellIndex));
    return _position(data.rowIndex, data.columnIndex);
  }

  static int _position(int row, int col) => row * 16384 + col;

  /// The cell operations used by wrapper.js `Cell` objects (see [JsCellOps]).
  JSObject get cellOps => _cellOps ??= createJSInteropWrapper(JsCellOps(_sheet));
  JSObject? _cellOps;

  List<CellValue?> _values(JSArray values) => [
        for (var i = 0; i < values.length; i++) jsToCellValue(values[i]),
      ];

  void appendRow(JSArray values) => _sheet.appendRow(_values(values));

  /// Appends several rows in one call, avoiding a JS↔Dart round trip per row.
  void appendRows(JSArray rows) {
    for (var i = 0; i < rows.length; i++) {
      _sheet.appendRow(_values(rows[i] as JSArray));
    }
  }

  /// Writes [values] into row [rowIndex] (0-based), from `startingColumn`.
  void insertRowIterables(JSArray values, int rowIndex, [JSAny? options]) {
    final o = optionsMap(options);
    _sheet.insertRowIterables(
      _values(values),
      rowIndex,
      startingColumn: o.integer('startingColumn') ?? 0,
      overwriteMergedCells: o.flag('overwriteMergedCells') ?? true,
    );
  }

  void insertRow(int rowIndex) => _sheet.insertRow(rowIndex);
  void removeRow(int rowIndex) => _sheet.removeRow(rowIndex);
  void insertColumn(int columnIndex) => _sheet.insertColumn(columnIndex);
  void removeColumn(int columnIndex) => _sheet.removeColumn(columnIndex);
  bool clearRow(int rowIndex) => _sheet.clearRow(rowIndex);

  /// Cell positions as in [cell], `null` where there is no cell.
  JSArray get rows {
    return plainJsArray(_sheet.rows.map((row) {
      return plainJsArray(row.map((cell) {
        return cell != null ? _position(cell.rowIndex, cell.columnIndex).toJS : null;
      }));
    }));
  }

  /// Values of a range such as `'A1:C3'` (or a single cell) as a 2D array.
  JSArray rangeValues(String range) => plainJsArray(_sheet
      .selectRangeValuesWithString(range)
      .map((row) => row == null ? JSArray() : plainJsArray(row.map(dartToJs))));

  /// Replaces [source] (a string or `RegExp`) with [target] in text cells;
  /// returns the number of replacements.
  int findAndReplace(JSAny source, String target, [JSAny? options]) {
    final Pattern pattern;
    if (source.isA<JSString>()) {
      pattern = (source as JSString).toDart;
    } else {
      final re = source as JSObject;
      final flags = (re.getProperty<JSString>('flags'.toJS)).toDart;
      pattern = RegExp(
        re.getProperty<JSString>('source'.toJS).toDart,
        caseSensitive: !flags.contains('i'),
        multiLine: flags.contains('m'),
        dotAll: flags.contains('s'),
        unicode: flags.contains('u'),
      );
    }
    final o = optionsMap(options);
    return _sheet.findAndReplace(
      pattern,
      target,
      first: o.integer('first') ?? -1,
      startingRow: o.integer('startingRow') ?? -1,
      endingRow: o.integer('endingRow') ?? -1,
      startingColumn: o.integer('startingColumn') ?? -1,
      endingColumn: o.integer('endingColumn') ?? -1,
    );
  }

  // ==================== DIMENSIONS & VISIBILITY ====================

  int? get frozenRows => _sheet.frozenRows;
  set frozenRows(int? value) => _sheet.frozenRows = value;

  int? get frozenColumns => _sheet.frozenColumns;
  set frozenColumns(int? value) => _sheet.frozenColumns = value;

  void setColumnHidden(int col, bool hidden) => _sheet.setColumnHidden(col, hidden);
  bool isColumnHidden(int col) => _sheet.isColumnHidden(col);
  void setRowHidden(int row, bool hidden) => _sheet.setRowHidden(row, hidden);
  bool isRowHidden(int row) => _sheet.isRowHidden(row);

  void setColumnWidth(int col, double width) => _sheet.setColumnWidth(col, width);
  double getColumnWidth(int col) => _sheet.getColumnWidth(col);
  void setRowHeight(int row, double height) => _sheet.setRowHeight(row, height);
  double getRowHeight(int row) => _sheet.getRowHeight(row);

  void setDefaultColumnWidth(double width) => _sheet.setDefaultColumnWidth(width);
  void setDefaultRowHeight(double height) => _sheet.setDefaultRowHeight(height);
  void setColumnAutoFit(int col) => _sheet.setColumnAutoFit(col);

  // ==================== MERGE ====================

  /// Merges the range; [customValue] replaces the merged content.
  void merge(String startCell, String endCell, [JSAny? customValue]) {
    _sheet.merge(
      CellIndex.indexByString(startCell),
      CellIndex.indexByString(endCell),
      customValue: jsToCellValue(customValue),
    );
  }

  /// [cellRef] is a merged range (`'A1:E1'`) or any cell inside one (`'A1'`).
  void unmerge(String cellRef) {
    final ref = cellRef.toUpperCase();
    if (ref.contains(':')) {
      _sheet.unMerge(ref);
      return;
    }
    final cell = CellIndex.indexByString(ref);
    for (final range in _sheet.spannedItems.toList()) {
      final parts = range.split(':');
      final start = CellIndex.indexByString(parts[0]);
      final end = CellIndex.indexByString(parts[1]);
      if (cell.rowIndex >= start.rowIndex &&
          cell.rowIndex <= end.rowIndex &&
          cell.columnIndex >= start.columnIndex &&
          cell.columnIndex <= end.columnIndex) {
        _sheet.unMerge(range);
      }
    }
  }

  JSArray<JSString> get spannedItems => plainJsArray(_sheet.spannedItems.map((s) => s.toJS));

  // ==================== AUTOFILTER ====================

  void setAutoFilter(String range) => _sheet.setAutoFilterByString(range);
  void clearAutoFilter() => _sheet.clearAutoFilter();
  bool get hasAutoFilter => _sheet.hasAutoFilter;

  /// Filter criteria for one column of the AutoFilter range.
  void addFilterColumn(JSObject options) => _sheet.addFilterColumn(parseFilterColumn(optionsMap(options)));

  JSObject? get autoFilter {
    final f = _sheet.autoFilter;
    if (f == null) return null;
    return jsObject({
      'ref': f.ref,
      'columns': [
        for (final c in f.filterColumns)
          {
            'column': c.colId,
            'values': c.filterValues,
            'blank': c.blank,
            'custom': [for (final r in c.customFilters) {'operator': r.operator.name, 'value': r.val}],
            'and': c.customFiltersAnd,
          }
      ],
    });
  }

  // ==================== PROTECTION ====================

  /// Protects the sheet. Each option set to `true` allows that action while
  /// the sheet is protected (like Excel's "Allow all users of this worksheet
  /// to" list); options left out stay blocked.
  void protect([String? password, JSObject? options]) {
    final p = _sheet.sheetProtection..sheet = true;
    if (password != null && password.isNotEmpty) {
      p.setPassword(password);
    }
    // The Dart flags (and the OOXML attributes) mean "this action is locked".
    final o = optionsMap(options);
    if (o.flag('objects') case final v?) p.objects = !v;
    if (o.flag('scenarios') case final v?) p.scenarios = !v;
    if (o.flag('formatCells') case final v?) p.formatCells = !v;
    if (o.flag('formatColumns') case final v?) p.formatColumns = !v;
    if (o.flag('formatRows') case final v?) p.formatRows = !v;
    if (o.flag('insertColumns') case final v?) p.insertColumns = !v;
    if (o.flag('insertRows') case final v?) p.insertRows = !v;
    if (o.flag('insertHyperlinks') case final v?) p.insertHyperlinks = !v;
    if (o.flag('deleteColumns') case final v?) p.deleteColumns = !v;
    if (o.flag('deleteRows') case final v?) p.deleteRows = !v;
    if (o.flag('selectLockedCells') case final v?) p.selectLockedCells = !v;
    if (o.flag('selectUnlockedCells') case final v?) p.selectUnlockedCells = !v;
    if (o.flag('sort') case final v?) p.sort = !v;
    if (o.flag('autoFilter') case final v?) p.autoFilter = !v;
    if (o.flag('pivotTables') case final v?) p.pivotTables = !v;
  }

  void unprotect() {
    _sheet.sheetProtection.sheet = false;
    _sheet.sheetProtection.password = null;
  }

  bool get isProtected => _sheet.sheetProtection.sheet;

  // ==================== GROUPING / OUTLINING ====================

  bool _collapsed(JSAny? options) => optionsMap(options).flag('collapsed') ?? false;

  void groupRows(int start, int end, [JSAny? options]) =>
      _sheet.groupRows(start, end, collapsed: _collapsed(options));
  void ungroupRows(int start, int end) => _sheet.ungroupRows(start, end);
  void groupColumns(int start, int end, [JSAny? options]) =>
      _sheet.groupColumns(start, end, collapsed: _collapsed(options));
  void ungroupColumns(int start, int end) => _sheet.ungroupColumns(start, end);

  void collapseRowGroup(int start, int end) => _sheet.collapseRowGroup(start, end);
  void expandRowGroup(int start, int end) => _sheet.expandRowGroup(start, end);
  void collapseColumnGroup(int start, int end) => _sheet.collapseColumnGroup(start, end);
  void expandColumnGroup(int start, int end) => _sheet.expandColumnGroup(start, end);
  void clearGrouping() => _sheet.clearGrouping();

  int getRowOutlineLevel(int row) => _sheet.getRowOutlineLevel(row);
  int getColumnOutlineLevel(int col) => _sheet.getColumnOutlineLevel(col);

  JSArray _groups(List<OutlineGroup> groups) => plainJsArray(groups.map((g) =>
      jsObject({'start': g.start, 'end': g.end, 'level': g.level, 'collapsed': g.collapsed})));

  JSArray get rowGroups => _groups(_sheet.rowGroups);
  JSArray get columnGroups => _groups(_sheet.columnGroups);

  JSObject get outlineSettings {
    final s = _sheet.outlineSettings;
    return jsObject({
      'summaryBelow': s.summaryBelow,
      'summaryRight': s.summaryRight,
      'showOutlineSymbols': s.showOutlineSymbols,
      'applyStyles': s.applyStyles,
    });
  }

  set outlineSettings(JSObject options) {
    final o = optionsMap(options);
    _sheet.outlineSettings = _sheet.outlineSettings.copyWith(
      summaryBelow: o.flag('summaryBelow'),
      summaryRight: o.flag('summaryRight'),
      showOutlineSymbols: o.flag('showOutlineSymbols'),
      applyStyles: o.flag('applyStyles'),
    );
  }

  // ==================== IMAGES ====================

  void addImage(JSUint8Array imageBytes, String format, int col, int row, int widthPixels, int heightPixels,
      [JSAny? options]) {
    final o = optionsMap(options);
    _sheet.addImage(ExcelImage(
      imageBytes: imageBytes.toDart,
      imageType: ExcelImageType.fromExtension(format),
      anchor: ImageAnchor.fromPixels(
        column: col,
        row: row,
        widthPixels: widthPixels,
        heightPixels: heightPixels,
        colOffsetPixels: o.integer('colOffset') ?? 0,
        rowOffsetPixels: o.integer('rowOffset') ?? 0,
      ),
    ));
  }

  // ==================== CHARTS ====================

  /// Accepts a config object or its JSON string.
  void addChart(JSAny config) => _sheet.addChart(parseChart(_options(config)));

  int get chartCount => _sheet.charts.length;

  // ==================== CONDITIONAL FORMATTING ====================

  /// Adds one rule object or an array of rules to [range].
  void addConditionalFormatting(String range, JSAny rules) {
    final value = rules.dartify();
    final list = value is List ? value : [value];
    _sheet.addConditionalFormatting(range, [
      for (final r in list) parseConditionalRule((r as Map).map((k, v) => MapEntry(k.toString(), v))),
    ]);
  }

  void clearConditionalFormatting() => _sheet.clearConditionalFormatting();

  JSArray get conditionalFormattings => plainJsArray(_sheet.conditionalFormattings.map((g) => jsObject({
        'range': g.sqref,
        'rules': [for (final r in g.rules) conditionalRuleInfo(r)],
      })));

  // ==================== DATA VALIDATION ====================

  void addDataValidation(String range, JSObject rule) =>
      _sheet.addDataValidation(range, parseDataValidation(optionsMap(rule)));

  void removeDataValidation(String range) => _sheet.removeDataValidation(range);
  void clearDataValidations() => _sheet.clearDataValidations();

  JSObject? getDataValidation(String cellRef) {
    final rule = _sheet.getDataValidation(CellIndex.indexByString(cellRef));
    return rule == null ? null : jsObject(dataValidationInfo(rule));
  }

  JSObject get dataValidations => jsObject({
        for (final e in _sheet.dataValidations.entries) e.key: dataValidationInfo(e.value),
      });

  // ==================== HYPERLINKS ====================

  /// Links every cell of [range] to the same target.
  void setHyperlinkRange(String range, JSObject target, [JSAny? options]) => _sheet.setHyperlinkRange(
        range,
        parseHyperlink(optionsMap(target)),
        styled: optionsMap(options).flag('styled') ?? true,
      );

  void clearHyperlinks() => _sheet.clearHyperlinks();

  JSObject get hyperlinks => jsObject({
        for (final e in _sheet.hyperlinks.entries) e.key: hyperlinkInfo(e.value),
      });

  // ==================== PAGE SETUP & PRINTING ====================

  /// Changes only the given options: `orientation`, `paperSize`, `scale`,
  /// `fitToWidth`, `fitToHeight`, `firstPageNumber`, `pageOrder`,
  /// `blackAndWhite`, `draft`, `cellComments`, `errors`, `copies`.
  void setPageSetup(JSObject options) => _sheet.pageSetup = parsePageSetup(_sheet.pageSetup, optionsMap(options));

  JSObject? get pageSetup => _sheet.pageSetup == null ? null : jsObject(pageSetupInfo(_sheet.pageSetup!));

  void clearPageSetup() => _sheet.clearPageSetup();

  /// `'normal' | 'wide' | 'narrow'` or `{left, right, top, bottom, header,
  /// footer, unit: 'in' | 'cm'}`.
  void setPageMargins(JSAny margins) => _sheet.pageMargins = parsePageMargins(margins.dartify()!);

  JSObject? get pageMargins => _sheet.pageMargins == null ? null : jsObject(pageMarginsInfo(_sheet.pageMargins!));

  void clearPageMargins() => _sheet.clearPageMargins();

  /// `{gridLines, headings, horizontalCentered, verticalCentered}`.
  void setPrintOptions(JSObject options) {
    final o = optionsMap(options);
    if (o.flag('gridLines') case final v?) _sheet.setPrintGridLines(v);
    if (o.flag('headings') case final v?) _sheet.setPrintHeadings(v);
    if (o.containsKey('horizontalCentered') || o.containsKey('verticalCentered')) {
      _sheet.setPrintCentered(
          horizontally: o.flag('horizontalCentered'), vertically: o.flag('verticalCentered'));
    }
  }

  JSObject? get printOptions {
    final p = _sheet.printOptions;
    return p == null
        ? null
        : jsObject({
            'gridLines': p.gridLines,
            'headings': p.headings,
            'horizontalCentered': p.horizontalCentered,
            'verticalCentered': p.verticalCentered,
          });
  }

  void clearPrintOptions() => _sheet.clearPrintOptions();

  /// `{header, footer, evenHeader, evenFooter, firstHeader, firstFooter,
  /// differentFirst, differentOddEven}` using Excel codes such as `&P`.
  void setHeaderFooter(JSObject options) =>
      _sheet.headerFooter = parseHeaderFooter(_sheet.headerFooter, optionsMap(options));

  JSObject? get headerFooter => _sheet.headerFooter == null ? null : jsObject(headerFooterInfo(_sheet.headerFooter!));

  void clearHeaderFooter() => _sheet.headerFooter = null;

  // ==================== TABLES ====================

  /// `addTable(range, name, columns?, options?)`: columns are names or
  /// `{name, totalsFunction, totalsLabel, ...}`; options are `style`,
  /// `showTotalsRow`, `showHeaderRow`, `showRowStripes`,
  /// `showColumnStripes`, `showFirstColumn`, `showLastColumn`,
  /// `showFilterButtons`.
  JSObject addTable(String range, String name, [JSAny? columns, JSAny? options]) {
    final o = optionsMap(options);
    final table = _sheet.addTable(
      range,
      name: name,
      columns: parseTableColumns(columns?.dartify()),
      style: o.containsKey('style') ? parseTableStyle(o.str('style')) : TableStyle.medium2,
      showHeaderRow: o.flag('showHeaderRow') ?? true,
      showTotalsRow: o.flag('showTotalsRow') ?? false,
      showRowStripes: o.flag('showRowStripes') ?? true,
      showColumnStripes: o.flag('showColumnStripes') ?? false,
      showFirstColumn: o.flag('showFirstColumn') ?? false,
      showLastColumn: o.flag('showLastColumn') ?? false,
      showFilterButtons: o.flag('showFilterButtons') ?? true,
    );
    return jsObject(tableInfo(table));
  }

  JSObject? getTable(String name) {
    final t = _sheet.getTable(name);
    return t == null ? null : jsObject(tableInfo(t));
  }

  JSArray get tables => plainJsArray(_sheet.tables.map((t) => jsObject(tableInfo(t))));

  /// Changes the options (and `columns`, `name`, `ref`) of a table.
  void updateTable(String name, JSObject options) {
    final table = _sheet.getTable(name);
    if (table == null) throw ArgumentError.value(name, 'name', 'no such table');
    final o = optionsMap(options);
    final style = o.containsKey('style') ? parseTableStyle(o.str('style')) : null;
    _sheet.updateTable(table.copyWith(
      name: o.str('name'),
      ref: o.str('ref'),
      columns: parseTableColumns(o['columns']),
      style: style,
      clearStyle: o.containsKey('style') && style == null,
      showHeaderRow: o.flag('showHeaderRow'),
      showTotalsRow: o.flag('showTotalsRow'),
      showRowStripes: o.flag('showRowStripes'),
      showColumnStripes: o.flag('showColumnStripes'),
      showFirstColumn: o.flag('showFirstColumn'),
      showLastColumn: o.flag('showLastColumn'),
      showFilterButtons: o.flag('showFilterButtons'),
    ));
  }

  void removeTable(String name) => _sheet.removeTable(name);

  JSArray tableRowsAsMaps(String name, [JSAny? options]) =>
      plainJsArray(_sheet.tableRowsAsMaps(name, mode: _mode(optionsMap(options))).map(dartToJs));

  JSObject appendTableRow(String name, JSArray values) =>
      jsObject(tableInfo(_sheet.appendTableRow(name, _values(values))));

  // ==================== PIVOT TABLES ====================

  void addPivotTable(JSObject options) => _sheet.addPivotTable(parsePivotTable(optionsMap(options)));

  int get pivotTableCount => _sheet.pivotTables.length;

  /// Recomputes the pivot table cells after their source data changed.
  void refreshPivotTables() => _sheet.refreshPivotTables();

  // ==================== EXPORT & IMPORT ====================

  ExportValueMode _mode(Map<String, Object?> o) =>
      o.str('mode') == 'displayText' ? ExportValueMode.displayText : ExportValueMode.typed;

  /// Rows below the header as objects keyed by the header cells. Options:
  /// `headerRow`, `mode: 'typed' | 'displayText'`, `skipEmptyRows`.
  JSArray rowsAsMaps([JSAny? options]) {
    final o = optionsMap(options);
    return plainJsArray(_sheet
        .rowsAsMaps(
          headerRow: o.integer('headerRow') ?? 0,
          mode: _mode(o),
          skipEmptyRows: o.flag('skipEmptyRows') ?? true,
        )
        .map(dartToJs));
  }

  /// Every row (header included) as an array of values.
  JSArray rowsAsValues([JSAny? options]) {
    final o = optionsMap(options);
    return plainJsArray(_sheet
        .rowsAsValues(mode: _mode(o), skipEmptyRows: o.flag('skipEmptyRows') ?? false)
        .map(dartToJs));
  }

  String toJson([JSAny? options]) {
    final o = optionsMap(options);
    return _sheet.toJson(
      headerRow: o.integer('headerRow') ?? 0,
      mode: _mode(o),
      skipEmptyRows: o.flag('skipEmptyRows') ?? true,
      indent: o.str('indent'),
    );
  }

  String toCsv([JSAny? options]) {
    final o = optionsMap(options);
    return _sheet.toCsv(
      separator: o.str('separator') ?? ',',
      lineTerminator: o.str('lineTerminator') ?? '\r\n',
      mode: o.str('mode') == 'typed' ? ExportValueMode.typed : ExportValueMode.displayText,
      skipEmptyRows: o.flag('skipEmptyRows') ?? false,
    );
  }

  /// Appends objects as rows, matching their keys with the header row
  /// (written first when the sheet is empty).
  void appendRowsFromMaps(JSArray rows, [JSAny? options]) {
    final o = optionsMap(options);
    final maps = [
      for (var i = 0; i < rows.length; i++)
        {
          for (final e in (rows[i].dartify() as Map).entries) e.key.toString(): dartToCellValue(e.value),
        },
    ];
    _sheet.appendRowsFromMaps(
      maps,
      headerRow: o.integer('headerRow') ?? 0,
      writeHeader: o.flag('writeHeader') ?? true,
    );
  }

  Map<String, Object?> _options(JSAny config) => config.isA<JSString>()
      ? optionsMap((config as JSString).toDart.toJSObjectFromJson())
      : optionsMap(config);
}

extension on String {
  /// Parses JSON into a JS value.
  JSAny toJSObjectFromJson() => _jsonParse(toJS);
}

@JS('JSON.parse')
external JSAny _jsonParse(JSString text);
