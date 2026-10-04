part of '../../excel_community.dart';

/// How cell values are converted by the export methods
/// ([SheetExport.rowsAsMaps], [SheetExport.rowsAsValues], ...).
enum ExportValueMode {
  /// Native Dart values: `String`, `int`, `double`, `bool`, `DateTime` (UTC)
  /// for dates, `Duration` for times. Formulas export their cached value
  /// (`null` when the file had none).
  typed,

  /// The text Excel displays for the cell, using its number format
  /// (see [Data.displayText]).
  displayText,
}

/// Exports worksheet data to Dart collections, JSON and CSV, and imports
/// rows from maps.
extension SheetExport on Sheet {
  /// Returns every row below [headerRow] as a map keyed by the header cells.
  ///
  /// Empty header cells are named after their column letter (`'C'`) and
  /// repeated headers get a numeric suffix (`'Name_2'`). Rows without any
  /// value are skipped unless [skipEmptyRows] is `false`.
  ///
  /// ```dart
  /// // Name | Age
  /// // Ana  | 31
  /// sheet.rowsAsMaps(); // [{'Name': 'Ana', 'Age': 31}]
  /// ```
  List<Map<String, Object?>> rowsAsMaps({
    int headerRow = 0,
    ExportValueMode mode = ExportValueMode.typed,
    bool skipEmptyRows = true,
  }) {
    if (headerRow < 0) {
      throw RangeError.value(headerRow, 'headerRow', 'must not be negative');
    }
    final headers = _exportHeaders(headerRow);
    final result = <Map<String, Object?>>[];
    for (var r = headerRow + 1; r < maxRows; r++) {
      if (skipEmptyRows && _isEmptyRow(r)) continue;
      result.add({
        for (var c = 0; c < headers.length; c++)
          headers[c]: _exportValue(_sheetData[r]?[c], mode),
      });
    }
    return result;
  }

  /// Returns the whole sheet as a grid of values (one list per row, each
  /// [maxColumns] long), converted according to [mode].
  List<List<Object?>> rowsAsValues({
    ExportValueMode mode = ExportValueMode.typed,
    bool skipEmptyRows = false,
  }) {
    return [
      for (var r = 0; r < maxRows; r++)
        if (!skipEmptyRows || !_isEmptyRow(r))
          [
            for (var c = 0; c < maxColumns; c++)
              _exportValue(_sheetData[r]?[c], mode),
          ],
    ];
  }

  /// Encodes [rowsAsMaps] as a JSON array of objects.
  ///
  /// Dates are written as `YYYY-MM-DD`, date-times as ISO 8601 and times as
  /// `HH:MM:SS`. Pass [indent] (e.g. `'  '`) for pretty-printed output.
  String toJson({
    int headerRow = 0,
    ExportValueMode mode = ExportValueMode.typed,
    bool skipEmptyRows = true,
    String? indent,
  }) {
    final rows = mode == ExportValueMode.typed
        ? _jsonRows(headerRow, skipEmptyRows)
        : rowsAsMaps(
            headerRow: headerRow, mode: mode, skipEmptyRows: skipEmptyRows);
    final encoder =
        indent == null ? const JsonEncoder() : JsonEncoder.withIndent(indent);
    return encoder.convert(rows);
  }

  /// Encodes the sheet as CSV (RFC 4180): fields containing the [separator],
  /// quotes or line breaks are quoted, and quotes are doubled.
  ///
  /// By default each cell is written as the text Excel displays.
  String toCsv({
    String separator = ',',
    String lineTerminator = '\r\n',
    ExportValueMode mode = ExportValueMode.displayText,
    bool skipEmptyRows = false,
  }) {
    String field(Object? value) {
      final text = switch (value) {
        null => '',
        DateTime d => d.toIso8601String(),
        Duration d => _formatDuration(d),
        _ => value.toString(),
      };
      final needsQuotes = text.contains(separator) ||
          text.contains('"') ||
          text.contains('\n') ||
          text.contains('\r');
      return needsQuotes ? '"${text.replaceAll('"', '""')}"' : text;
    }

    return rowsAsValues(mode: mode, skipEmptyRows: skipEmptyRows)
        .map((row) => row.map(field).join(separator))
        .join(lineTerminator);
  }

  /// Appends [rows] below the existing data, one row per map.
  ///
  /// When the sheet is empty, a header row with the keys is written first
  /// (unless [writeHeader] is `false`). With an existing header row
  /// ([headerRow]), values are placed under the matching header; keys
  /// without a matching header are added as new columns.
  ///
  /// Values are converted to cell values: `String`, `int`, `double`, `bool`,
  /// `DateTime` (a [DateCellValue] when it has no time part), `Duration`
  /// ([TimeCellValue]) or any [CellValue]. Other objects use `toString()`.
  void appendRowsFromMaps(
    List<Map<String, Object?>> rows, {
    int headerRow = 0,
    bool writeHeader = true,
  }) {
    if (rows.isEmpty) return;

    final headers = <String>[];
    if (maxRows > headerRow) {
      final row = _sheetData[headerRow];
      for (var c = 0; c < maxColumns; c++) {
        final value = row?[c]?.value;
        headers.add(value == null ? '' : _headerText(row![c]!));
      }
    } else if (writeHeader) {
      // Keys in first-seen order across all maps.
      for (final map in rows) {
        for (final key in map.keys) {
          if (!headers.contains(key)) headers.add(key);
        }
      }
      insertRowIterables(
          [for (final h in headers) TextCellValue(h)], headerRow);
    }

    for (final map in rows) {
      for (final key in map.keys) {
        if (!headers.contains(key)) {
          headers.add(key);
          if (writeHeader && maxRows > headerRow) {
            updateCell(
                CellIndex.indexByColumnRow(
                    columnIndex: headers.length - 1, rowIndex: headerRow),
                TextCellValue(key));
          }
        }
      }
      final values = List<CellValue?>.filled(headers.length, null);
      map.forEach((key, value) {
        values[headers.indexOf(key)] = _toCellValue(value);
      });
      appendRow(values);
    }
  }

  List<String> _exportHeaders(int headerRow) {
    final headers = <String>[];
    final seen = <String, int>{};
    for (var c = 0; c < maxColumns; c++) {
      final cell = _sheetData[headerRow]?[c];
      var name = cell?.value == null ? '' : _headerText(cell!).trim();
      if (name.isEmpty) name = getColumnAlphabet(c);
      final count = (seen[name] ?? 0) + 1;
      seen[name] = count;
      headers.add(count == 1 ? name : '${name}_$count');
    }
    return headers;
  }

  String _headerText(Data cell) {
    final value = cell.value;
    return value is TextCellValue ? value.toString() : cell.displayText;
  }

  bool _isEmptyRow(int rowIndex) {
    final row = _sheetData[rowIndex];
    if (row == null) return true;
    return row.values.every((cell) => cell.value == null);
  }

  Object? _exportValue(Data? cell, ExportValueMode mode) {
    if (cell == null || cell.value == null) return null;
    return mode == ExportValueMode.displayText
        ? cell.displayText
        : _typedValue(cell.value);
  }

  List<Map<String, Object?>> _jsonRows(int headerRow, bool skipEmptyRows) {
    final headers = _exportHeaders(headerRow);
    return [
      for (var r = headerRow + 1; r < maxRows; r++)
        if (!skipEmptyRows || !_isEmptyRow(r))
          {
            for (var c = 0; c < headers.length; c++)
              headers[c]: _jsonValue(_sheetData[r]?[c]?.value),
          },
    ];
  }
}

/// Native Dart value of a cell value (see [ExportValueMode.typed]).
Object? _typedValue(CellValue? value) => switch (value) {
      null => null,
      TextCellValue v => v.toString(),
      IntCellValue v => v.value,
      DoubleCellValue v => v.value,
      BoolCellValue v => v.value,
      DateCellValue v => v.asDateTimeUtc(),
      DateTimeCellValue v => v.asDateTimeUtc(),
      TimeCellValue v => v.asDuration(),
      FormulaCellValue v => _typedValue(v.cachedValue),
    };

/// JSON-safe value of a cell value.
Object? _jsonValue(CellValue? value) => switch (value) {
      null => null,
      DoubleCellValue v => v.value.isFinite ? v.value : null,
      DateCellValue v => '${v.year.toString().padLeft(4, '0')}-'
          '${v.month.toString().padLeft(2, '0')}-'
          '${v.day.toString().padLeft(2, '0')}',
      DateTimeCellValue v => v.asDateTimeUtc().toIso8601String(),
      TimeCellValue v => _formatDuration(v.asDuration()),
      FormulaCellValue v => _jsonValue(v.cachedValue),
      _ => _typedValue(value),
    };

String _formatDuration(Duration d) {
  String two(int n) => n.toString().padLeft(2, '0');
  final base = '${two(d.inHours)}:${two(d.inMinutes.remainder(60))}:'
      '${two(d.inSeconds.remainder(60))}';
  final ms = d.inMilliseconds.remainder(1000);
  return ms == 0 ? base : '$base.${ms.toString().padLeft(3, '0')}';
}

CellValue? _toCellValue(Object? value) => switch (value) {
      null => null,
      CellValue v => v,
      String v => TextCellValue(v),
      int v => IntCellValue(v),
      double v => DoubleCellValue(v),
      bool v => BoolCellValue(v),
      DateTime v => v.hour == 0 &&
              v.minute == 0 &&
              v.second == 0 &&
              v.millisecond == 0 &&
              v.microsecond == 0
          ? DateCellValue(year: v.year, month: v.month, day: v.day)
          : DateTimeCellValue.fromDateTime(v),
      Duration v => TimeCellValue.fromDuration(v),
      _ => TextCellValue(value.toString()),
    };

/// Exports every worksheet of a workbook.
extension ExcelExport on Excel {
  /// [SheetExport.rowsAsMaps] for every sheet, keyed by sheet name.
  Map<String, List<Map<String, Object?>>> toMaps({
    int headerRow = 0,
    ExportValueMode mode = ExportValueMode.typed,
    bool skipEmptyRows = true,
  }) {
    return {
      for (final entry in tables.entries)
        entry.key: entry.value.rowsAsMaps(
            headerRow: headerRow, mode: mode, skipEmptyRows: skipEmptyRows),
    };
  }

  /// A JSON object with one array per sheet: `{"Sheet1": [{...}], ...}`.
  String toJson({
    int headerRow = 0,
    ExportValueMode mode = ExportValueMode.typed,
    bool skipEmptyRows = true,
    String? indent,
  }) {
    final sheets = <String, Object?>{
      for (final entry in tables.entries)
        entry.key: jsonDecode(entry.value.toJson(
            headerRow: headerRow, mode: mode, skipEmptyRows: skipEmptyRows)),
    };
    final encoder =
        indent == null ? const JsonEncoder() : JsonEncoder.withIndent(indent);
    return encoder.convert(sheets);
  }
}
