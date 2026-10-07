import 'dart:js_interop';
import 'package:excel_community/excel_community.dart';
import 'js_convert.dart';
import 'js_options.dart';

/// The cell operations of one sheet, addressed by 0-based row and column.
///
/// Building a JS object with `createJSInteropWrapper` costs about a
/// microsecond per member, so instead of one wrapper per cell, the `Cell`
/// class in wrapper.js holds a position and calls this object, which is
/// created once per sheet.
@JSExport()
class JsCellOps {
  final Sheet _sheet;
  JsCellOps(this._sheet);

  Data _data(int row, int col) =>
      _sheet.cell(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: col));

  String displayText(int row, int col) => _data(row, col).displayText;

  String? getComment(int row, int col) => _data(row, col).comment;
  void setComment(int row, int col, String? val) => _data(row, col).comment = val;

  String? getFormula(int row, int col) {
    final v = _data(row, col).value;
    return v is FormulaCellValue ? v.formula : null;
  }

  void setFormula(int row, int col, String? formula) {
    if (formula != null) {
      _data(row, col).setFormula(formula);
    }
  }

  /// The result Excel cached for a formula cell when the file was saved.
  JSAny? cachedValue(int row, int col) {
    final v = _data(row, col).value;
    return v is FormulaCellValue ? cellValueToJs(v.cachedValue) : null;
  }

  String type(int row, int col) => switch (_data(row, col).value) {
        TextCellValue() => 'string',
        IntCellValue() => 'int',
        DoubleCellValue() => 'double',
        BoolCellValue() => 'bool',
        DateCellValue() => 'date',
        DateTimeCellValue() => 'datetime',
        TimeCellValue() => 'time',
        FormulaCellValue() => 'formula',
        null => 'null',
      };

  JSAny? getValue(int row, int col) => cellValueToJs(_data(row, col).value);

  void setValue(int row, int col, JSAny? val) => _data(row, col).value = jsToCellValue(val);

  /// A JS `Date` for date and date-time cells, otherwise `null`.
  JSAny? dateValue(int row, int col) => switch (_data(row, col).value) {
        DateCellValue v => jsDate(v.year, v.month, v.day),
        DateTimeCellValue v => jsDate(v.year, v.month, v.day, v.hour, v.minute, v.second, v.millisecond),
        _ => null,
      };

  /// Stores a time of day, from `'HH:MM[:SS]'` or hour/minute/second numbers.
  void setTime(int row, int col, JSAny time, [int? minute, int? second]) {
    _data(row, col).value = time.isA<JSString>()
        ? TimeCellValue.fromDuration(parseTime((time as JSString).toDart))
        : TimeCellValue(
            hour: (time as JSNumber).toDartInt,
            minute: minute ?? 0,
            second: second ?? 0,
          );
  }

  // ==================== STYLING ====================

  /// Changes only the given style options; the rest of the style is kept.
  void setStyle(int row, int col, JSObject options) {
    final data = _data(row, col);
    data.cellStyle = applyStyleOptions(data.cellStyle, optionsMap(options));
  }

  /// Removes all formatting from the cell.
  void resetStyle(int row, int col) {
    _data(row, col).cellStyle = CellStyle();
  }

  JSObject? style(int row, int col) {
    final s = _data(row, col).cellStyle;
    return s == null ? null : jsObject(styleInfo(s));
  }

  // ==================== HYPERLINKS ====================

  /// `setHyperlink(url, tooltip?, text?)` or `setHyperlink({url | email |
  /// sheet+cell | location, tooltip, display}, {text, styled})`.
  void setHyperlink(int row, int col, JSAny target, [JSAny? tooltipOrOptions, String? display]) {
    final index = CellIndex.indexByColumnRow(rowIndex: row, columnIndex: col);
    if (target.isA<JSString>()) {
      final tooltip = tooltipOrOptions.isA<JSString>() ? (tooltipOrOptions as JSString).toDart : null;
      _sheet.setHyperlink(
        index,
        Hyperlink.url((target as JSString).toDart, tooltip: tooltip, display: display),
        text: display,
      );
      return;
    }
    final options = optionsMap(tooltipOrOptions);
    _sheet.setHyperlink(
      index,
      parseHyperlink(optionsMap(target)),
      text: options.str('text'),
      styled: options.flag('styled') ?? true,
    );
  }

  JSObject? getHyperlink(int row, int col) {
    final link = _sheet.getHyperlink(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: col));
    return link == null ? null : jsObject(hyperlinkInfo(link));
  }

  void removeHyperlink(int row, int col) {
    _sheet.removeHyperlink(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: col));
  }

  // ==================== DATA VALIDATION ====================

  JSObject? dataValidation(int row, int col) {
    final rule = _sheet.getDataValidation(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: col));
    return rule == null ? null : jsObject(dataValidationInfo(rule));
  }

  /// Whether [value] passes this cell's validation (`null` when it depends
  /// on formulas or other cells).
  bool? validates(int row, int col, JSAny? value) => _sheet
          .getDataValidation(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: col))
          ?.accepts(jsToCellValue(value)) ??
      true;
}
