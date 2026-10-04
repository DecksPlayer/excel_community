import 'dart:convert';
import 'dart:js_interop';
import 'dart:js_interop_unsafe';
import 'package:excel_community/excel_community.dart';
import 'js_cell.dart';

@JSExport()
class JsSheet {
  final Sheet _sheet;
  JsSheet(this._sheet);

  String get name => _sheet.sheetName;
  int get maxRows => _sheet.maxRows;
  int get maxColumns => _sheet.maxColumns;

  bool get rightToLeft => _sheet.isRTL;
  set rightToLeft(bool val) => _sheet.isRTL = val;

  String? get tabColor => _sheet.tabColor?.rgb ?? _sheet.tabColor?.color?.colorHex;
  set tabColor(String? hex) {
    if (hex == null || hex.isEmpty) {
      _sheet.clearTabColor();
    } else {
      _sheet.setTabColorHex(hex);
    }
  }

  JSObject cell(String cellIndex) {
    return createJSInteropWrapper(JsCell(_sheet.cell(CellIndex.indexByString(cellIndex)), _sheet));
  }

  void appendRow(JSArray values) {
    final list = <CellValue?>[];
    for (int i = 0; i < values.length; i++) {
      final v = values[i];
      if (v == null) {
        list.add(null);
      } else if (v.isA<JSBoolean>()) {
        list.add(BoolCellValue((v as JSBoolean).toDart));
      } else if (v.isA<JSNumber>()) {
        final d = (v as JSNumber).toDartDouble;
        if (d == d.roundToDouble() && !d.isInfinite && !d.isNaN) {
          list.add(IntCellValue(d.toInt()));
        } else {
          list.add(DoubleCellValue(d));
        }
      } else if (v.isA<JSString>()) {
        final s = (v as JSString).toDart;
        list.add(s.startsWith('=') ? FormulaCellValue(s) : TextCellValue(s));
      }
    }
    _sheet.appendRow(list);
  }

  JSArray get rows {
    return _sheet.rows.map((row) {
      return row.map((cell) {
        return cell != null ? createJSInteropWrapper(JsCell(cell, _sheet)) : null;
      }).toList().toJS;
    }).toList().toJS;
  }

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

  void merge(String startCell, String endCell) {
    _sheet.merge(CellIndex.indexByString(startCell), CellIndex.indexByString(endCell));
  }

  void unmerge(String cellRef) {
    _sheet.unMerge(cellRef);
  }

  JSArray<JSString> get spannedItems =>
      _sheet.spannedItems.map((s) => s.toJS).toList().toJS;

  void setAutoFilter(String range) => _sheet.setAutoFilterByString(range);
  void clearAutoFilter() => _sheet.clearAutoFilter();
  bool get hasAutoFilter => _sheet.hasAutoFilter;

  // ==================== PROTECTION ====================

  void protect([String? password, JSObject? options]) {
    _sheet.sheetProtection.sheet = true;
    if (password != null && password.isNotEmpty) {
      _sheet.sheetProtection.setPassword(password);
    }
    if (options != null) {
      final fc = options['formatCells'];
      if (fc != null && fc.isA<JSBoolean>()) _sheet.sheetProtection.formatCells = (fc as JSBoolean).toDart;
      final fcol = options['formatColumns'];
      if (fcol != null && fcol.isA<JSBoolean>()) _sheet.sheetProtection.formatColumns = (fcol as JSBoolean).toDart;
      final frow = options['formatRows'];
      if (frow != null && frow.isA<JSBoolean>()) _sheet.sheetProtection.formatRows = (frow as JSBoolean).toDart;
      final ic = options['insertColumns'];
      if (ic != null && ic.isA<JSBoolean>()) _sheet.sheetProtection.insertColumns = (ic as JSBoolean).toDart;
      final ir = options['insertRows'];
      if (ir != null && ir.isA<JSBoolean>()) _sheet.sheetProtection.insertRows = (ir as JSBoolean).toDart;
      final dc = options['deleteColumns'];
      if (dc != null && dc.isA<JSBoolean>()) _sheet.sheetProtection.deleteColumns = (dc as JSBoolean).toDart;
      final dr = options['deleteRows'];
      if (dr != null && dr.isA<JSBoolean>()) _sheet.sheetProtection.deleteRows = (dr as JSBoolean).toDart;
      final af = options['autoFilter'];
      if (af != null && af.isA<JSBoolean>()) _sheet.sheetProtection.autoFilter = (af as JSBoolean).toDart;
      final s = options['sort'];
      if (s != null && s.isA<JSBoolean>()) _sheet.sheetProtection.sort = (s as JSBoolean).toDart;
    }
  }

  void unprotect() {
    _sheet.sheetProtection.sheet = false;
    _sheet.sheetProtection.password = null;
  }

  // ==================== GROUPING / OUTLINING ====================

  void groupRows(int start, int end) => _sheet.groupRows(start, end);
  void ungroupRows(int start, int end) => _sheet.ungroupRows(start, end);
  void groupColumns(int start, int end) => _sheet.groupColumns(start, end);
  void ungroupColumns(int start, int end) => _sheet.ungroupColumns(start, end);

  // ==================== IMAGES ====================

  void addImage(JSUint8Array imageBytes, String format, int col, int row, int widthPixels, int heightPixels) {
    _sheet.addImage(ExcelImage(
      imageBytes: imageBytes.toDart,
      imageType: ExcelImageType.fromExtension(format),
      anchor: ImageAnchor.fromPixels(
        column: col,
        row: row,
        widthPixels: widthPixels,
        heightPixels: heightPixels,
      ),
    ));
  }

  // ==================== CHARTS ====================

  void addChart(String chartConfigJson) {
    final Map<String, dynamic> cfg = jsonDecode(chartConfigJson);
    final title = cfg['title']?.toString() ?? 'Chart';
    final type = cfg['type']?.toString().toLowerCase() ?? 'column';
    final showLegend = cfg['showLegend'] != false;

    final anchorMap = cfg['anchor'] as Map<String, dynamic>?;
    final int fromCol = anchorMap?['fromCol'] ?? 0;
    final int fromRow = anchorMap?['fromRow'] ?? 0;
    final int toCol = anchorMap?['toCol'] ?? (fromCol + 8);
    final int toRow = anchorMap?['toRow'] ?? (fromRow + 15);
    final anchor = ChartAnchor(
      fromColumn: fromCol,
      fromRow: fromRow,
      toColumn: toCol,
      toRow: toRow,
    );

    final seriesList = <ChartSeries>[];
    final rawSeries = cfg['series'] as List<dynamic>? ?? [];
    for (final s in rawSeries) {
      if (s is Map) {
        ChartSeriesStyle? style;
        if (s['colorHex'] != null) {
          style = ChartSeriesStyle(fillColor: ExcelColor.fromHexString(s['colorHex'].toString()));
        }
        seriesList.add(ChartSeries(
          name: s['name']?.toString() ?? '',
          categoriesRange: s['categoriesRange']?.toString() ?? '',
          valuesRange: s['valuesRange']?.toString() ?? '',
          style: style,
        ));
      }
    }

    ChartGrouping grouping = ChartGrouping.clustered;
    final gStr = cfg['grouping']?.toString().toLowerCase();
    if (gStr == 'stacked') {
      grouping = ChartGrouping.stacked;
    } else if (gStr == 'percentstacked') {
      grouping = ChartGrouping.percentStacked;
    }

    Chart chart;
    switch (type) {
      case 'bar':
        chart = ColumnChart(title: title, series: seriesList, anchor: anchor, showLegend: showLegend, isVertical: false, grouping: grouping);
        break;
      case 'line':
        chart = LineChart(
          title: title,
          series: seriesList,
          anchor: anchor,
          showLegend: showLegend,
          grouping: grouping,
          showMarkers: cfg['showMarkers'] != false,
          smooth: cfg['smooth'] == true,
        );
        break;
      case 'pie':
        chart = PieChart(title: title, series: seriesList, anchor: anchor, showLegend: showLegend);
        break;
      case 'area':
        chart = AreaChart(title: title, series: seriesList, anchor: anchor, showLegend: showLegend, grouping: grouping);
        break;
      case 'scatter':
        chart = ScatterChart(title: title, series: seriesList, anchor: anchor, showLegend: showLegend);
        break;
      case 'radar':
        chart = RadarChart(title: title, series: seriesList, anchor: anchor, showLegend: showLegend);
        break;
      case 'column':
      default:
        chart = ColumnChart(title: title, series: seriesList, anchor: anchor, showLegend: showLegend, isVertical: true, grouping: grouping);
        break;
    }
    _sheet.addChart(chart);
  }

  // ==================== TABLES ====================

  void addTable(String range, String name, [String? columnsJson]) {
    List<TableColumn>? columns;
    if (columnsJson != null) {
      final List<dynamic> list = jsonDecode(columnsJson);
      columns = list.map((c) => TableColumn(c.toString())).toList();
    }
    _sheet.addTable(range, name: name, columns: columns);
  }
}
