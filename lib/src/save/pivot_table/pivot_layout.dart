part of '../../../excel_community.dart';

/// One row or column of a pivot table: a data line (an item path, plus the
/// data field on the column axis when there are several values), a subtotal
/// (`default`) or a grand total (`grand`).
class _PivotAxisEntry {
  final String type; // 'data' | 'default' | 'grand'
  final List<int> path;
  final int dataIndex;

  /// How many leading items equal the previous entry's (`r` attribute).
  int repeated = 0;

  _PivotAxisEntry(this.type, this.path, {this.dataIndex = 0});
}

/// A source value in the pivot cache.
class _PivotCacheValue {
  final String kind; // 's' | 'n' | 'd' | 'm'
  final String text;
  final num? number;

  const _PivotCacheValue(this.kind, this.text, [this.number]);

  static const missing = _PivotCacheValue('m', '');

  String get key => '$kind:$text';

  CellValue? get cellValue => switch (kind) {
        's' => TextCellValue(text),
        'n' => number == number!.roundToDouble() && number!.abs() < 1e15
            ? IntCellValue(number!.toInt())
            : DoubleCellValue(number!.toDouble()),
        'd' => () {
            final d = DateTime.parse(text);
            return d.hour == 0 && d.minute == 0 && d.second == 0
                ? DateCellValue(year: d.year, month: d.month, day: d.day)
                : DateTimeCellValue.fromDateTime(d);
          }(),
        _ => TextCellValue('(blank)'),
      };
}

/// Computes a pivot table the way Excel lays it out (compact rows with
/// subtotals on top, tabular columns, grand totals), so the file holds the
/// finished table and its cache instead of an empty shell that only Excel
/// can fill by refreshing.
class _PivotLayout {
  final PivotTable pivot;
  final List<String> headers;

  /// Source records: one cache value per field.
  final List<List<_PivotCacheValue>> records;

  final List<int> rowFields;
  final List<int> colFields;
  final List<int> dataFieldIndexes;
  final List<PivotTableValue> dataValues;
  late final List<String> dataNames;

  /// Distinct values of each axis field in order of appearance, and the
  /// item index of every record for that field.
  final Map<int, List<_PivotCacheValue>> sharedItems = {};
  final Map<int, List<int>> recordItems = {};

  late final List<_PivotAxisEntry> rowEntries;
  late final List<_PivotAxisEntry> colEntries;

  _PivotLayout._(this.pivot, this.headers, this.records, this.rowFields,
      this.colFields, this.dataFieldIndexes, this.dataValues) {
    for (final field in {...rowFields, ...colFields}) {
      final keys = <String, int>{};
      final items = <_PivotCacheValue>[];
      recordItems[field] = [
        for (final record in records)
          keys.putIfAbsent(record[field].key, () {
            items.add(record[field]);
            return items.length - 1;
          }),
      ];
      sharedItems[field] = items;
    }
    dataNames = _uniqueDataNames();
    rowEntries = _buildRowEntries();
    colEntries = _buildColEntries();
  }

  /// Returns `null` when the source sheet or headers cannot be read.
  static _PivotLayout? build(PivotTable pivot, Sheet source) {
    final _CellRect rect;
    try {
      rect = _CellRect.parse(pivot.sourceRange);
    } catch (_) {
      return null;
    }
    CellValue? valueAt(int row, int col) => source._sheetData[row]?[col]?.value;

    final headers = [
      for (var c = rect.left; c <= rect.right; c++)
        valueAt(rect.top, c)?.toString() ?? getCellId(c, rect.top),
    ];
    if (headers.isEmpty) return null;

    final records = <List<_PivotCacheValue>>[];
    for (var r = rect.top + 1; r <= rect.bottom; r++) {
      final record = [
        for (var c = rect.left; c <= rect.right; c++) _cacheValue(valueAt(r, c)),
      ];
      if (record.any((v) => v.kind != 'm')) records.add(record);
    }

    List<int> indexes(List<String> names) =>
        [for (final n in names) if (headers.indexOf(n) case final i when i != -1) i];
    final values = [for (final v in pivot.values) if (headers.contains(v.field)) v];
    return _PivotLayout._(
      pivot,
      headers,
      records,
      indexes(pivot.rows),
      indexes(pivot.columns),
      [for (final v in values) headers.indexOf(v.field)],
      values,
    );
  }

  static _PivotCacheValue _cacheValue(CellValue? value) => switch (value) {
        null => _PivotCacheValue.missing,
        TextCellValue v => () {
            final text = v.value.text ?? v.toString();
            return text.isEmpty ? _PivotCacheValue.missing : _PivotCacheValue('s', text);
          }(),
        IntCellValue v => _PivotCacheValue('n', '${v.value}', v.value),
        DoubleCellValue v => _PivotCacheValue('n', _formatNumber(v.value), v.value),
        BoolCellValue v => _PivotCacheValue('s', v.value ? 'TRUE' : 'FALSE'),
        DateCellValue v => _PivotCacheValue('d', _isoDate(v.year, v.month, v.day, 0, 0, 0)),
        DateTimeCellValue v =>
          _PivotCacheValue('d', _isoDate(v.year, v.month, v.day, v.hour, v.minute, v.second)),
        TimeCellValue v => () {
            final fraction = (v.hour * 3600 + v.minute * 60 + v.second) / 86400;
            return _PivotCacheValue('n', _formatNumber(fraction), fraction);
          }(),
        FormulaCellValue v => _cacheValue(v.cachedValue),
      };

  static String _isoDate(int y, int m, int d, int h, int min, int s) {
    String two(int n) => n.toString().padLeft(2, '0');
    return '${y.toString().padLeft(4, '0')}-${two(m)}-${two(d)}T${two(h)}:${two(min)}:${two(s)}';
  }

  static String _formatNumber(num n) =>
      n == n.roundToDouble() && n.abs() < 1e15 ? n.toInt().toString() : n.toString();

  // ==================== NAMES ====================

  static String functionLabel(PivotValueFunction f) => switch (f) {
        PivotValueFunction.sum => 'Sum',
        PivotValueFunction.count => 'Count',
        PivotValueFunction.average => 'Average',
        PivotValueFunction.max => 'Max',
        PivotValueFunction.min => 'Min',
        PivotValueFunction.product => 'Product',
        PivotValueFunction.countNums => 'Count',
        PivotValueFunction.stdDev => 'StdDev',
        PivotValueFunction.stdDevp => 'StdDevp',
        PivotValueFunction.varVal => 'Var',
        PivotValueFunction.varp => 'Varp',
      };

  /// Excel rejects data field names that repeat or equal a source field name.
  List<String> _uniqueDataNames() {
    final used = <String>{for (final h in headers) h.toLowerCase()};
    return [
      for (final v in dataValues)
        () {
          var name = v.customName ?? '${functionLabel(v.function)} of ${v.field}';
          if (used.contains(name.toLowerCase())) {
            final base = name;
            var n = 2;
            name = v.customName != null ? '$base ' : '$base$n';
            while (used.contains(name.toLowerCase())) {
              name = '$base${n++}';
            }
          }
          used.add(name.toLowerCase());
          return name;
        }(),
    ];
  }

  // ==================== AXES ====================

  /// Item paths present in the data for [fields], in item order, depth first.
  void _walk(List<int> fields, List<int> recordIndexes, List<int> prefix,
      void Function(List<int> path, List<int> recordIndexes) visit, {void Function(List<int> path)? after}) {
    final depth = prefix.length;
    if (depth == fields.length) return;
    final field = fields[depth];
    final byItem = <int, List<int>>{};
    for (final r in recordIndexes) {
      byItem.putIfAbsent(recordItems[field]![r], () => []).add(r);
    }
    for (final item in byItem.keys.toList()..sort()) {
      final path = [...prefix, item];
      visit(path, byItem[item]!);
      _walk(fields, byItem[item]!, path, visit, after: after);
      after?.call(path);
    }
  }

  List<int> get _allRecords => [for (var i = 0; i < records.length; i++) i];

  void _setRepeated(List<_PivotAxisEntry> entries) {
    List<int>? previous;
    for (final e in entries) {
      if (e.type == 'data' && previous != null) {
        var r = 0;
        while (r < e.path.length && r < previous.length && e.path[r] == previous[r]) {
          r++;
        }
        e.repeated = r;
      }
      previous = e.path;
    }
  }

  /// Compact form: each item on its own row, parents (with their subtotal)
  /// above their children, then the grand total.
  List<_PivotAxisEntry> _buildRowEntries() {
    if (rowFields.isEmpty) return [_PivotAxisEntry('data', const [])];
    final entries = <_PivotAxisEntry>[];
    _walk(rowFields, _allRecords, const [], (path, _) => entries.add(_PivotAxisEntry('data', path)));
    entries.add(_PivotAxisEntry('grand', const [0]));
    _setRepeated(entries);
    return entries;
  }

  /// Tabular form: leaf columns (one per value when there are several), a
  /// subtotal column after each outer item, then the grand totals.
  List<_PivotAxisEntry> _buildColEntries() {
    final k = dataValues.length;
    final entries = <_PivotAxisEntry>[];
    if (colFields.isEmpty) {
      if (k > 1) {
        for (var d = 0; d < k; d++) {
          entries.add(_PivotAxisEntry('data', [d], dataIndex: d));
        }
      } else {
        entries.add(_PivotAxisEntry('data', const []));
      }
      _setRepeated(entries);
      return entries;
    }
    _walk(colFields, _allRecords, const [], (path, _) {
      if (path.length < colFields.length) return;
      if (k > 1) {
        for (var d = 0; d < k; d++) {
          entries.add(_PivotAxisEntry('data', [...path, d], dataIndex: d));
        }
      } else {
        entries.add(_PivotAxisEntry('data', path));
      }
    }, after: (path) {
      if (path.length == colFields.length) return;
      for (var d = 0; d < k; d++) {
        entries.add(_PivotAxisEntry('default', path, dataIndex: d));
      }
    });
    for (var d = 0; d < k; d++) {
      entries.add(_PivotAxisEntry('grand', const [0], dataIndex: d));
    }
    _setRepeated(entries);
    return entries;
  }

  // ==================== VALUES ====================

  /// Item path of the real axis fields that an entry filters on.
  List<int> _filterPath(_PivotAxisEntry e, int fieldCount) =>
      e.type == 'grand' ? const [] : (e.path.length > fieldCount ? e.path.sublist(0, fieldCount) : e.path);

  bool _matches(int record, List<int> fields, List<int> path) {
    for (var i = 0; i < path.length; i++) {
      if (recordItems[fields[i]]![record] != path[i]) return false;
    }
    return true;
  }

  /// The aggregated value of data field [d] for a row and column entry, or
  /// `null` when Excel would leave the cell empty.
  num? valueAt(_PivotAxisEntry row, _PivotAxisEntry col) {
    final d = dataValues.length > 1 ? col.dataIndex : 0;
    if (d >= dataValues.length) return null;
    final rowPath = row.type == 'grand' ? const <int>[] : row.path;
    final colPath = _filterPath(col, colFields.length);
    final field = dataFieldIndexes[d];
    final matched = [
      for (var r = 0; r < records.length; r++)
        if (_matches(r, rowFields, rowPath) && _matches(r, colFields, colPath)) records[r][field],
    ];
    if (matched.isEmpty) return null;
    final numbers = [for (final v in matched) if (v.kind == 'n') v.number!];
    final n = numbers.length;
    num mean() => numbers.reduce((a, b) => a + b) / n;
    num sumSq() => numbers.fold<num>(0, (s, x) => s + (x - mean()) * (x - mean()));
    return switch (dataValues[d].function) {
      PivotValueFunction.sum => numbers.fold<num>(0, (a, b) => a + b),
      PivotValueFunction.count => matched.where((v) => v.kind != 'm').length,
      PivotValueFunction.countNums => n,
      PivotValueFunction.average => n == 0 ? null : mean(),
      PivotValueFunction.max => n == 0 ? 0 : numbers.reduce(max),
      PivotValueFunction.min => n == 0 ? 0 : numbers.reduce(min),
      PivotValueFunction.product => n == 0 ? 0 : numbers.reduce((a, b) => a * b),
      PivotValueFunction.stdDev => n < 2 ? null : sqrt(sumSq() / (n - 1)),
      PivotValueFunction.stdDevp => n < 1 ? null : sqrt(sumSq() / n),
      PivotValueFunction.varVal => n < 2 ? null : sumSq() / (n - 1),
      PivotValueFunction.varp => n < 1 ? null : sumSq() / n,
    };
  }

  // ==================== GEOMETRY ====================

  /// Column levels: the column fields plus "Values" when there are several.
  int get colLevels => colFields.length + (dataValues.length > 1 ? 1 : 0);
  int get headerRows => colLevels == 0 ? 1 : 1 + colLevels;
  int get firstDataCol => rowFields.isNotEmpty || colFields.isNotEmpty ? 1 : 0;
  int get rowCount => headerRows + rowEntries.length;
  int get columnCount => firstDataCol + colEntries.length;

  String get locationRef {
    final top = pivot.targetCell.rowIndex;
    final left = pivot.targetCell.columnIndex;
    return '${getCellId(left, top)}:${getCellId(left + columnCount - 1, top + rowCount - 1)}';
  }

  String _itemLabel(int field, int item) {
    final v = sharedItems[field]![item];
    return v.kind == 'm' ? '(blank)' : v.text;
  }

  CellValue _itemValue(int field, int item) =>
      sharedItems[field]![item].cellValue ?? TextCellValue(_itemLabel(field, item));

  /// The cells Excel shows for this pivot table, relative to its top-left
  /// cell, as (row, column) → value. Numbers are flagged to use the General
  /// format.
  Map<(int, int), (CellValue, bool)> cells() {
    final out = <(int, int), (CellValue, bool)>{};
    void text(int r, int c, String s) => out[(r, c)] = (TextCellValue(s), false);
    final k = dataValues.length;
    final m = colFields.length;
    final fdc = firstDataCol;

    // Header rows.
    if (colLevels == 0) {
      if (fdc == 1) text(0, 0, 'Row Labels');
      if (k > 0) text(0, fdc, dataNames[0]);
    } else {
      if (rowFields.isNotEmpty && k == 1) text(0, 0, dataNames[0]);
      text(0, fdc, m > 0 ? 'Column Labels' : 'Values');
      if (rowFields.isNotEmpty) text(colLevels, 0, 'Row Labels');
      for (var ci = 0; ci < colEntries.length; ci++) {
        final e = colEntries[ci];
        final col = fdc + ci;
        if (e.type == 'grand') {
          text(1, col, k > 1 ? 'Total ${dataNames[e.dataIndex]}' : 'Grand Total');
        } else if (e.type == 'default') {
          final item = _itemLabel(colFields[e.path.length - 1], e.path.last);
          text(e.path.length, col, k > 1 ? '$item ${dataNames[e.dataIndex]}' : '$item Total');
        } else {
          for (var level = e.repeated; level < e.path.length; level++) {
            if (level < m) {
              out[(1 + level, col)] = (_itemValue(colFields[level], e.path[level]), false);
            } else {
              text(1 + level, col, dataNames[e.path[level]]);
            }
          }
        }
      }
    }

    // Data rows.
    for (var ri = 0; ri < rowEntries.length; ri++) {
      final row = rowEntries[ri];
      final r = headerRows + ri;
      if (fdc == 1) {
        if (rowFields.isEmpty) {
          if (k == 1) text(r, 0, dataNames[0]);
        } else if (row.type == 'grand') {
          text(r, 0, 'Grand Total');
        } else {
          out[(r, 0)] = (_itemValue(rowFields[row.path.length - 1], row.path.last), false);
        }
      }
      for (var ci = 0; ci < colEntries.length; ci++) {
        final value = valueAt(row, colEntries[ci]);
        if (value != null) {
          out[(r, fdc + ci)] = (
            value == value.roundToDouble() && value.abs() < 1e15
                ? IntCellValue(value.toInt())
                : DoubleCellValue(value.toDouble()),
            true,
          );
        }
      }
    }
    return out;
  }
}
