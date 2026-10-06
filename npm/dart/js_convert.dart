import 'dart:js_interop';
import 'dart:js_interop_unsafe';
import 'package:excel_community/excel_community.dart';

/// Conversions between JS values and the Dart model, shared by the bindings.

/// Converts a JS value into a [CellValue]: booleans, numbers, strings
/// (`=...` becomes a formula) and `Date` objects. Anything else is `null`.
CellValue? jsToCellValue(JSAny? val) {
  if (val == null || val.isUndefinedOrNull) return null;
  if (val.isA<JSBoolean>()) return BoolCellValue((val as JSBoolean).toDart);
  if (val.isA<JSNumber>()) return numToCellValue((val as JSNumber).toDartDouble);
  if (val.isA<JSString>()) {
    final s = (val as JSString).toDart;
    return s.startsWith('=') ? FormulaCellValue(s) : TextCellValue(s);
  }
  return dartToCellValue(val.dartify());
}

/// Like [jsToCellValue] for a value already converted with `dartify()`.
CellValue? dartToCellValue(Object? value) => switch (value) {
      null => null,
      bool v => BoolCellValue(v),
      num v => numToCellValue(v.toDouble()),
      String v => v.startsWith('=') ? FormulaCellValue(v) : TextCellValue(v),
      DateTime v => dateToCellValue(v),
      _ => null,
    };

CellValue numToCellValue(double d) =>
    d == d.roundToDouble() && d.isFinite ? IntCellValue(d.toInt()) : DoubleCellValue(d);

/// Excel stores wall-clock time, so a JS `Date` keeps its local time; dates
/// at midnight become date-only cells.
CellValue dateToCellValue(DateTime date) {
  final d = date.toLocal();
  final midnight = d.hour == 0 && d.minute == 0 && d.second == 0 && d.millisecond == 0;
  return midnight
      ? DateCellValue(year: d.year, month: d.month, day: d.day)
      : DateTimeCellValue.fromDateTime(d);
}

/// Parses `'HH:MM'` or `'HH:MM:SS'` into a [Duration].
Duration parseTime(String text) {
  final parts = text.trim().split(':').map(int.parse).toList();
  if (parts.length < 2 || parts.length > 3) {
    throw ArgumentError.value(text, 'time', "expected 'HH:MM' or 'HH:MM:SS'");
  }
  return Duration(hours: parts[0], minutes: parts[1], seconds: parts.length > 2 ? parts[2] : 0);
}

String _two(int n) => n.toString().padLeft(2, '0');

String formatDate(int year, int month, int day) =>
    '${year.toString().padLeft(4, '0')}-${_two(month)}-${_two(day)}';

String formatTime(int hour, int minute, int second) => '${_two(hour)}:${_two(minute)}:${_two(second)}';

String formatDuration(Duration d) =>
    formatTime(d.inHours, d.inMinutes.remainder(60), d.inSeconds.remainder(60));

/// A JS `Date` with the given local wall-clock time.
JSObject jsDate(int year, int month, int day, [int hour = 0, int minute = 0, int second = 0, int ms = 0]) =>
    (globalContext['Date'] as JSFunction).callAsConstructorVarArgs<JSObject>([
      year.toJS,
      (month - 1).toJS,
      day.toJS,
      hour.toJS,
      minute.toJS,
      second.toJS,
      ms.toJS,
    ]);

/// The JS value of a cell: dates as `'YYYY-MM-DD'`, date-times as ISO
/// strings, times as `'HH:MM:SS'` and formulas as their text.
JSAny? cellValueToJs(CellValue? v) => switch (v) {
      null => null,
      TextCellValue() => (v.value.text ?? v.toString()).toJS,
      IntCellValue() => v.value.toJS,
      DoubleCellValue() => v.value.toJS,
      BoolCellValue() => v.value.toJS,
      DateCellValue() => formatDate(v.year, v.month, v.day).toJS,
      DateTimeCellValue() => v.asDateTimeLocal().toIso8601String().toJS,
      TimeCellValue() => formatTime(v.hour, v.minute, v.second).toJS,
      FormulaCellValue() => v.formula.toJS,
    };

/// Converts exported Dart values (maps, lists, primitives, `DateTime`,
/// `Duration`, cell values) into plain JS values. Dates become JS `Date`s
/// with the same wall-clock time.
JSAny? dartToJs(Object? value) => switch (value) {
      null => null,
      String v => v.toJS,
      bool v => v.toJS,
      num v => v.toJS,
      DateTime v => jsDate(v.year, v.month, v.day, v.hour, v.minute, v.second, v.millisecond),
      Duration v => formatDuration(v).toJS,
      CellValue v => cellValueToJs(v),
      Map v => _jsObject(v),
      Iterable v => plainJsArray(v.map(dartToJs)),
      _ => value.toString().toJS,
    };

JSObject _jsObject(Map map) {
  final obj = JSObject();
  map.forEach((key, value) => obj[key.toString()] = dartToJs(value));
  return obj;
}

/// Builds a plain JS array (no dart2js type metadata attached).
JSArray<T> plainJsArray<T extends JSAny?>(Iterable<T> items) {
  final list = items.toList(growable: false);
  final arr = JSArray<T>.withLength(list.length);
  for (var i = 0; i < list.length; i++) {
    arr[i] = list[i];
  }
  return arr;
}

/// A JS object built from [entries], skipping `null` values.
JSObject jsObject(Map<String, Object?> entries) {
  final obj = JSObject();
  entries.forEach((key, value) {
    if (value != null) obj[key] = dartToJs(value);
  });
  return obj;
}

/// Reads a JS options object (or its JSON string) as a Dart map.
Map<String, Object?> optionsMap(JSAny? options) {
  if (options == null || options.isUndefinedOrNull) return const {};
  final value = options.dartify();
  if (value is Map) return value.map((k, v) => MapEntry(k.toString(), v));
  throw ArgumentError('Expected an options object');
}

/// Typed accessors for option maps produced by [optionsMap] / `dartify()`.
extension OptionReaders on Map<String, Object?> {
  String? str(String key) => this[key]?.toString();
  bool? flag(String key) => this[key] is bool ? this[key] as bool : null;
  int? integer(String key) => (this[key] as num?)?.toInt();
  double? number(String key) => (this[key] as num?)?.toDouble();
  Map<String, Object?>? obj(String key) {
    final v = this[key];
    return v is Map ? v.map((k, v) => MapEntry(k.toString(), v)) : null;
  }

  List<Object?>? list(String key) => this[key] is List ? this[key] as List : null;
}

ExcelColor parseColor(String hex) => ExcelColor.fromHexString(hex.startsWith('#') ? hex : '#$hex');

/// Finds an enum value by name, ignoring case, `-`, `_` and spaces.
T enumByName<T extends Enum>(List<T> values, String? name, T fallback) {
  if (name == null) return fallback;
  String norm(String s) => s.toLowerCase().replaceAll(RegExp(r'[-_ ]'), '');
  final wanted = norm(name);
  for (final v in values) {
    if (norm(v.name) == wanted) return v;
  }
  throw ArgumentError.value(name, 'name', 'expected one of ${values.map((v) => v.name).join(', ')}');
}

T? enumByNameOrNull<T extends Enum>(List<T> values, String? name) =>
    name == null ? null : enumByName(values, name, values.first);
