import 'dart:js_interop';
import 'package:excel_community/excel_community.dart';
import 'js_convert.dart';
import 'js_options.dart';

@JSExport()
class JsCell {
  final Data _data;
  final Sheet _sheet;
  JsCell(this._data, this._sheet);

  String get cellId => _data.cellIndex.cellId;
  int get row => _data.rowIndex;
  int get col => _data.columnIndex;
  String get displayText => _data.displayText;

  String? get comment => _data.comment;
  set comment(String? val) => _data.comment = val;

  String? get formula {
    final v = _data.value;
    return v is FormulaCellValue ? v.formula : null;
  }

  set formula(String? val) {
    if (val != null) {
      _data.setFormula(val);
    }
  }

  void setFormula(String formula) {
    _data.setFormula(formula);
  }

  /// The result Excel cached for a formula cell when the file was saved.
  JSAny? get cachedValue {
    final v = _data.value;
    return v is FormulaCellValue ? cellValueToJs(v.cachedValue) : null;
  }

  String get type => switch (_data.value) {
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

  JSAny? get value => cellValueToJs(_data.value);

  set value(JSAny? val) => _data.value = jsToCellValue(val);

  /// A JS `Date` for date and date-time cells, otherwise `null`.
  JSAny? get dateValue => switch (_data.value) {
        DateCellValue v => jsDate(v.year, v.month, v.day),
        DateTimeCellValue v => jsDate(v.year, v.month, v.day, v.hour, v.minute, v.second, v.millisecond),
        _ => null,
      };

  /// Stores a time of day, from `'HH:MM[:SS]'` or hour/minute/second numbers.
  void setTime(JSAny time, [int? minute, int? second]) {
    _data.value = time.isA<JSString>()
        ? TimeCellValue.fromDuration(parseTime((time as JSString).toDart))
        : TimeCellValue(
            hour: (time as JSNumber).toDartInt,
            minute: minute ?? 0,
            second: second ?? 0,
          );
  }

  // ==================== STYLING ====================

  /// Changes only the given style options; the rest of the style is kept.
  void setStyle(JSObject options) {
    _data.cellStyle = applyStyleOptions(_data.cellStyle, optionsMap(options));
  }

  /// Removes all formatting from the cell.
  void resetStyle() {
    _data.cellStyle = CellStyle();
  }

  JSObject? get style {
    final s = _data.cellStyle;
    return s == null ? null : jsObject(styleInfo(s));
  }

  // ==================== HYPERLINKS ====================

  /// `setHyperlink(url, tooltip?, text?)` or `setHyperlink({url | email |
  /// sheet+cell | location, tooltip, display}, {text, styled})`.
  void setHyperlink(JSAny target, [JSAny? tooltipOrOptions, String? display]) {
    if (target.isA<JSString>()) {
      final tooltip = tooltipOrOptions.isA<JSString>() ? (tooltipOrOptions as JSString).toDart : null;
      _sheet.setHyperlink(
        _data.cellIndex,
        Hyperlink.url((target as JSString).toDart, tooltip: tooltip, display: display),
        text: display,
      );
      return;
    }
    final options = optionsMap(tooltipOrOptions);
    _sheet.setHyperlink(
      _data.cellIndex,
      parseHyperlink(optionsMap(target)),
      text: options.str('text'),
      styled: options.flag('styled') ?? true,
    );
  }

  JSObject? getHyperlink() {
    final link = _sheet.getHyperlink(_data.cellIndex);
    return link == null ? null : jsObject(hyperlinkInfo(link));
  }

  void removeHyperlink() {
    _sheet.removeHyperlink(_data.cellIndex);
  }

  // ==================== DATA VALIDATION ====================

  JSObject? get dataValidation {
    final rule = _sheet.getDataValidation(_data.cellIndex);
    return rule == null ? null : jsObject(dataValidationInfo(rule));
  }

  /// Whether [value] passes this cell's validation (`null` when it depends
  /// on formulas or other cells).
  bool? validates(JSAny? value) =>
      _sheet.getDataValidation(_data.cellIndex)?.accepts(jsToCellValue(value)) ?? true;
}
