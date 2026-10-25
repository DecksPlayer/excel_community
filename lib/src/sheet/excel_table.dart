part of '../../excel_community.dart';

/// Aggregate shown in a table's totals row (`totalsRowFunction`).
enum TableTotalsFunction {
  none('none', null),
  sum('sum', 109),
  average('average', 101),
  count('count', 103),
  countNumbers('countNums', 102),
  min('min', 105),
  max('max', 104),
  stdDev('stdDev', 107),
  variance('var', 110),

  /// Uses [TableColumn.totalsRowFormula].
  custom('custom', null);

  final String xmlValue;

  /// `SUBTOTAL` function number (ignores hidden rows), when built in.
  final int? subtotalCode;

  const TableTotalsFunction(this.xmlValue, this.subtotalCode);

  static TableTotalsFunction fromXmlValue(String? value) =>
      values.firstWhere((f) => f.xmlValue == value, orElse: () => none);
}

/// Built-in table style (`<tableStyleInfo name="...">`).
class TableStyle extends Equatable {
  final String name;

  /// Any style name, e.g. a custom style defined in the workbook.
  const TableStyle.named(this.name);

  /// `TableStyleLight1` … `TableStyleLight21`.
  factory TableStyle.light(int number) => _builtIn('Light', number, 21);

  /// `TableStyleMedium1` … `TableStyleMedium28`.
  factory TableStyle.medium(int number) => _builtIn('Medium', number, 28);

  /// `TableStyleDark1` … `TableStyleDark11`.
  factory TableStyle.dark(int number) => _builtIn('Dark', number, 11);

  /// Excel's default table style (blue, banded rows).
  static const medium2 = TableStyle.named('TableStyleMedium2');
  static const light9 = TableStyle.named('TableStyleLight9');
  static const dark1 = TableStyle.named('TableStyleDark1');

  static TableStyle _builtIn(String kind, int number, int count) {
    RangeError.checkValueInInterval(number, 1, count, 'number');
    return TableStyle.named('TableStyle$kind$number');
  }

  @override
  List<Object?> get props => [name];

  @override
  String toString() => name;
}

/// A column of an [ExcelTable] (`<tableColumn>`).
class TableColumn extends Equatable {
  /// Header text; must be unique within the table.
  final String name;

  /// Aggregate shown in the totals row.
  final TableTotalsFunction totalsFunction;

  /// Text shown in the totals row (typically on the first column).
  final String? totalsLabel;

  /// Formula for [TableTotalsFunction.custom] (without `=`).
  final String? totalsRowFormula;

  /// Formula filled down the column, as stored in the file (kept when a
  /// file is re-saved).
  final String? calculatedColumnFormula;

  const TableColumn(
    this.name, {
    this.totalsFunction = TableTotalsFunction.none,
    this.totalsLabel,
    this.totalsRowFormula,
    this.calculatedColumnFormula,
  });

  TableColumn copyWith({
    String? name,
    TableTotalsFunction? totalsFunction,
    String? totalsLabel,
    String? totalsRowFormula,
    String? calculatedColumnFormula,
  }) {
    return TableColumn(
      name ?? this.name,
      totalsFunction: totalsFunction ?? this.totalsFunction,
      totalsLabel: totalsLabel ?? this.totalsLabel,
      totalsRowFormula: totalsRowFormula ?? this.totalsRowFormula,
      calculatedColumnFormula:
          calculatedColumnFormula ?? this.calculatedColumnFormula,
    );
  }

  @override
  List<Object?> get props =>
      [name, totalsFunction, totalsLabel, totalsRowFormula, calculatedColumnFormula];
}

/// An Excel table ("Format as Table", `xl/tables/tableN.xml`).
///
/// ```dart
/// sheet.addTable('A1:D20',
///     name: 'Sales',
///     style: TableStyle.medium(9),
///     showTotalsRow: true,
///     columns: const [
///       TableColumn('Region', totalsLabel: 'Total'),
///       TableColumn('Units', totalsFunction: TableTotalsFunction.sum),
///       TableColumn('Price'),
///       TableColumn('Revenue', totalsFunction: TableTotalsFunction.sum),
///     ]);
/// ```
class ExcelTable extends Equatable {
  /// Name used in formulas (`Sales[Revenue]`); unique in the workbook.
  final String name;

  /// Range including the header and totals rows, e.g. `A1:D20`.
  final String ref;

  final List<TableColumn> columns;

  /// `null` for no style.
  final TableStyle? style;

  final bool showHeaderRow;
  final bool showTotalsRow;
  final bool showRowStripes;
  final bool showColumnStripes;
  final bool showFirstColumn;
  final bool showLastColumn;

  /// Whether the header row shows filter buttons.
  final bool showFilterButtons;

  const ExcelTable({
    required this.name,
    required this.ref,
    required this.columns,
    this.style = TableStyle.medium2,
    this.showHeaderRow = true,
    this.showTotalsRow = false,
    this.showRowStripes = true,
    this.showColumnStripes = false,
    this.showFirstColumn = false,
    this.showLastColumn = false,
    this.showFilterButtons = true,
  });

  _CellRect get _rect => _CellRect.parse(ref);

  /// Row index (0-based) of the header row, or `null` without one.
  int? get headerRowIndex => showHeaderRow ? _rect.top : null;

  /// Row index of the totals row, or `null` without one.
  int? get totalsRowIndex => showTotalsRow ? _rect.bottom : null;

  /// First and last data row indexes (0-based, inclusive).
  int get firstDataRow => _rect.top + (showHeaderRow ? 1 : 0);
  int get lastDataRow => _rect.bottom - (showTotalsRow ? 1 : 0);

  /// Number of data rows.
  int get dataRowCount => lastDataRow - firstDataRow + 1;

  /// Column index (0-based) of the column named [columnName], or `-1`.
  int columnIndexOf(String columnName) {
    final i = columns.indexWhere(
        (c) => c.name.toLowerCase() == columnName.toLowerCase());
    return i < 0 ? -1 : _rect.left + i;
  }

  /// Structured reference to a column, e.g. `Sales[Revenue]`.
  String columnReference(String columnName) =>
      '$name[${_escapeStructuredName(columnName)}]';

  ExcelTable copyWith({
    String? name,
    String? ref,
    List<TableColumn>? columns,
    TableStyle? style,
    bool clearStyle = false,
    bool? showHeaderRow,
    bool? showTotalsRow,
    bool? showRowStripes,
    bool? showColumnStripes,
    bool? showFirstColumn,
    bool? showLastColumn,
    bool? showFilterButtons,
  }) {
    return ExcelTable(
      name: name ?? this.name,
      ref: ref ?? this.ref,
      columns: columns ?? this.columns,
      style: clearStyle ? null : (style ?? this.style),
      showHeaderRow: showHeaderRow ?? this.showHeaderRow,
      showTotalsRow: showTotalsRow ?? this.showTotalsRow,
      showRowStripes: showRowStripes ?? this.showRowStripes,
      showColumnStripes: showColumnStripes ?? this.showColumnStripes,
      showFirstColumn: showFirstColumn ?? this.showFirstColumn,
      showLastColumn: showLastColumn ?? this.showLastColumn,
      showFilterButtons: showFilterButtons ?? this.showFilterButtons,
    );
  }

  /// The `xl/tables/tableN.xml` part for this table.
  String _toXmlString(int id) {
    final rect = _rect;
    final sb = StringBuffer(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'id="$id" name="${_escapeXml(name)}" displayName="${_escapeXml(name)}" '
        'ref="${rect.ref}"');
    if (!showHeaderRow) sb.write(' headerRowCount="0"');
    if (showTotalsRow) {
      sb.write(' totalsRowCount="1"');
    } else {
      sb.write(' totalsRowShown="0"');
    }
    sb.write('>');
    if (showHeaderRow && showFilterButtons) {
      final filterRect = _CellRect(rect.top, rect.left, lastDataRow, rect.right);
      sb.write('<autoFilter ref="${filterRect.ref}"/>');
    }
    sb.write('<tableColumns count="${columns.length}">');
    for (var i = 0; i < columns.length; i++) {
      final c = columns[i];
      sb.write('<tableColumn id="${i + 1}" name="${_escapeXml(c.name)}"');
      if (c.totalsFunction != TableTotalsFunction.none) {
        sb.write(' totalsRowFunction="${c.totalsFunction.xmlValue}"');
      }
      if (c.totalsLabel != null) {
        sb.write(' totalsRowLabel="${_escapeXml(c.totalsLabel!)}"');
      }
      if (c.calculatedColumnFormula == null && c.totalsRowFormula == null) {
        sb.write('/>');
        continue;
      }
      sb.write('>');
      if (c.calculatedColumnFormula != null) {
        sb.write('<calculatedColumnFormula>'
            '${_escapeXml(c.calculatedColumnFormula!)}</calculatedColumnFormula>');
      }
      if (c.totalsRowFormula != null) {
        sb.write('<totalsRowFormula>${_escapeXml(c.totalsRowFormula!)}</totalsRowFormula>');
      }
      sb.write('</tableColumn>');
    }
    sb.write('</tableColumns>');
    sb.write('<tableStyleInfo');
    if (style != null) sb.write(' name="${_escapeXml(style!.name)}"');
    sb.write(' showFirstColumn="${showFirstColumn ? 1 : 0}"'
        ' showLastColumn="${showLastColumn ? 1 : 0}"'
        ' showRowStripes="${showRowStripes ? 1 : 0}"'
        ' showColumnStripes="${showColumnStripes ? 1 : 0}"/>');
    sb.write('</table>');
    return sb.toString();
  }

  /// Parses a `<table>` part.
  static ExcelTable? _fromXml(XmlDocument document) {
    final root = document.rootElement;
    final ref = root.getAttribute('ref');
    final name = root.getAttribute('displayName') ?? root.getAttribute('name');
    if (ref == null || name == null) return null;
    bool flag(XmlElement? e, String attr, bool fallback) {
      final v = e?.getAttribute(attr);
      return v == null ? fallback : (v == '1' || v == 'true');
    }

    final styleInfo = root.findElements('tableStyleInfo').firstOrNull;
    final styleName = styleInfo?.getAttribute('name');
    return ExcelTable(
      name: name,
      ref: _CellRect.parse(ref).ref,
      columns: [
        for (final col in root.findAllElements('tableColumn'))
          TableColumn(
            col.getAttribute('name') ?? '',
            totalsFunction:
                TableTotalsFunction.fromXmlValue(col.getAttribute('totalsRowFunction')),
            totalsLabel: col.getAttribute('totalsRowLabel'),
            totalsRowFormula: col.findElements('totalsRowFormula').firstOrNull?.innerText,
            calculatedColumnFormula:
                col.findElements('calculatedColumnFormula').firstOrNull?.innerText,
          ),
      ],
      style: styleName == null ? null : TableStyle.named(styleName),
      showHeaderRow: (int.tryParse(root.getAttribute('headerRowCount') ?? '') ?? 1) > 0,
      showTotalsRow: (int.tryParse(root.getAttribute('totalsRowCount') ?? '') ?? 0) > 0,
      showRowStripes: flag(styleInfo, 'showRowStripes', true),
      showColumnStripes: flag(styleInfo, 'showColumnStripes', false),
      showFirstColumn: flag(styleInfo, 'showFirstColumn', false),
      showLastColumn: flag(styleInfo, 'showLastColumn', false),
      showFilterButtons: root.findElements('autoFilter').isNotEmpty,
    );
  }

  @override
  List<Object?> get props => [
        name,
        ref,
        columns,
        style,
        showHeaderRow,
        showTotalsRow,
        showRowStripes,
        showColumnStripes,
        showFirstColumn,
        showLastColumn,
        showFilterButtons,
      ];
}

/// Escapes `[`, `]`, `#` and `'` for use inside a structured reference.
String _escapeStructuredName(String name) =>
    name.replaceAllMapped(RegExp(r"[\[\]#']"), (m) => "'${m[0]}");

final _tableNamePattern = RegExp(r'^[A-Za-z_\\][A-Za-z0-9_.\\]*$');
final _cellLikeName = RegExp(r'^([A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*|[RrCc])$');

/// Excel tables of a worksheet.
extension SheetTables on Sheet {
  /// Tables of this worksheet, in creation order.
  List<ExcelTable> get tables => List.unmodifiable(_tables);

  bool get hasTables => _tables.isNotEmpty;

  /// The table named [name] (case-insensitive), if it is on this sheet.
  ExcelTable? getTable(String name) {
    final lower = name.toLowerCase();
    for (final t in _tables) {
      if (t.name.toLowerCase() == lower) return t;
    }
    return null;
  }

  /// The table containing [cellIndex], if any.
  ExcelTable? tableAt(CellIndex cellIndex) {
    for (final t in _tables) {
      if (t._rect.contains(cellIndex)) return t;
    }
    return null;
  }

  /// Formats [range] (`'A1:D20'`, header and totals rows included) as an
  /// Excel table and returns it.
  ///
  /// Without [columns], the names come from the header row (empty or
  /// repeated headers become `Column1`, `Name2`, ...). The header cells are
  /// set to the column names, and the totals row (if any) gets each
  /// column's label or `SUBTOTAL` formula.
  ///
  /// Throws an [ArgumentError] for an invalid or duplicate [name], a column
  /// count that does not match the range, too few rows, a range that
  /// overlaps another table, merged cells or the sheet's AutoFilter, or a
  /// totals row ([showTotalsRow]) whose cells already hold data.
  ExcelTable addTable(
    String range, {
    required String name,
    List<TableColumn>? columns,
    TableStyle? style = TableStyle.medium2,
    bool showHeaderRow = true,
    bool showTotalsRow = false,
    bool showRowStripes = true,
    bool showColumnStripes = false,
    bool showFirstColumn = false,
    bool showLastColumn = false,
    bool showFilterButtons = true,
  }) {
    final rect = _CellRect.parse(range);
    _validateTableName(name);
    final width = rect.right - rect.left + 1;
    final minRows = 1 + (showHeaderRow ? 1 : 0) + (showTotalsRow ? 1 : 0);
    if (rect.bottom - rect.top + 1 < minRows) {
      throw ArgumentError.value(range, 'range',
          'needs at least $minRows rows (header, one data row and totals)');
    }
    for (final other in _tables) {
      if (other._rect.intersects(rect)) {
        throw ArgumentError.value(range, 'range', 'overlaps table "${other.name}"');
      }
    }
    for (final span in _spanList) {
      if (span == null) continue;
      final merged = _CellRect(span.rowSpanStart, span.columnSpanStart,
          span.rowSpanEnd, span.columnSpanEnd);
      if (merged.intersects(rect)) {
        throw ArgumentError.value(range, 'range', 'contains merged cells');
      }
    }
    final filter = _autoFilter;
    if (filter != null && _CellRect.parse(filter.ref).intersects(rect)) {
      throw ArgumentError.value(range, 'range',
          "overlaps the sheet's AutoFilter; clear it first");
    }

    final resolvedColumns = columns ?? _columnsFromHeader(rect, showHeaderRow);
    if (resolvedColumns.length != width) {
      throw ArgumentError.value(columns, 'columns',
          'must have $width entries for $range');
    }
    final names = <String>{};
    for (final c in resolvedColumns) {
      if (c.name.trim().isEmpty || !names.add(c.name.toLowerCase())) {
        throw ArgumentError.value(c.name, 'columns', 'names must be unique and not empty');
      }
    }

    final table = ExcelTable(
      name: name,
      ref: rect.ref,
      columns: resolvedColumns,
      style: style,
      showHeaderRow: showHeaderRow,
      showTotalsRow: showTotalsRow,
      showRowStripes: showRowStripes,
      showColumnStripes: showColumnStripes,
      showFirstColumn: showFirstColumn,
      showLastColumn: showLastColumn,
      showFilterButtons: showFilterButtons,
    );
    _checkTotalsRowIsFree(table);
    _tables.add(table);
    _syncTableCells(table);
    return table;
  }

  /// Removes the table named [name]; its cells keep their values.
  void removeTable(String name) {
    final lower = name.toLowerCase();
    _tables.removeWhere((t) => t.name.toLowerCase() == lower);
  }

  /// Replaces the table that has the same name as [table] (e.g. a
  /// `copyWith` of it) and refreshes its header and totals cells.
  ///
  /// Like Excel, turning the totals row on without changing the range adds
  /// a row below the table (shifting the rows under it when they have
  /// data), and turning it off removes that row from the table.
  void updateTable(ExcelTable table) {
    final i = _tables.indexWhere((t) => t.name.toLowerCase() == table.name.toLowerCase());
    if (i < 0) throw ArgumentError.value(table.name, 'table', 'not found');
    final old = _tables[i];
    if (table.ref == old.ref && table.showTotalsRow != old.showTotalsRow) {
      final rect = old._rect;
      if (table.showTotalsRow) {
        final below = rect.bottom + 1;
        final occupied = [
          for (var c = rect.left; c <= rect.right; c++) _sheetData[below]?[c]?.value,
        ].any((v) => v != null);
        if (occupied) insertRow(below);
        table = table.copyWith(ref: _CellRect(rect.top, rect.left, below, rect.right).ref);
      } else {
        for (var c = rect.left; c <= rect.right; c++) {
          if (_sheetData[rect.bottom]?[c] != null) {
            updateCell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: rect.bottom), null);
          }
        }
        table = table.copyWith(ref: _CellRect(rect.top, rect.left, rect.bottom - 1, rect.right).ref);
      }
    } else if (table.showTotalsRow &&
        (!old.showTotalsRow || table.totalsRowIndex != old.totalsRowIndex)) {
      _checkTotalsRowIsFree(table);
    }
    _tables[i] = table;
    _syncTableCells(table);
  }

  /// The data rows of table [name] as maps keyed by column name.
  List<Map<String, Object?>> tableRowsAsMaps(String name,
      {ExportValueMode mode = ExportValueMode.typed}) {
    final table = getTable(name);
    if (table == null) throw ArgumentError.value(name, 'name', 'no such table');
    final left = table._rect.left;
    return [
      for (var r = table.firstDataRow; r <= table.lastDataRow; r++)
        {
          for (var i = 0; i < table.columns.length; i++)
            table.columns[i].name: _exportValue(_sheetData[r]?[left + i], mode),
        },
    ];
  }

  /// Adds a data row at the end of table [name] (above its totals row,
  /// which moves down) and returns the updated table.
  ExcelTable appendTableRow(String name, List<CellValue?> values) {
    final table = getTable(name);
    if (table == null) throw ArgumentError.value(name, 'name', 'no such table');
    if (values.length > table.columns.length) {
      throw ArgumentError.value(values, 'values', 'more values than columns');
    }
    final rect = table._rect;
    final int row;
    if (table.showTotalsRow) {
      row = rect.bottom;
      insertRow(row); // grows the table and moves the totals row down
    } else {
      row = rect.bottom + 1;
      final grown = table.copyWith(
          ref: _CellRect(rect.top, rect.left, row, rect.right).ref);
      _tables[_tables.indexOf(table)] = grown;
    }
    for (var i = 0; i < values.length; i++) {
      updateCell(
          CellIndex.indexByColumnRow(columnIndex: rect.left + i, rowIndex: row),
          values[i]);
    }
    final updated = getTable(name)!;
    _syncTableCells(updated);
    return updated;
  }

  void _validateTableName(String name) {
    if (name.isEmpty ||
        name.length > 255 ||
        !_tableNamePattern.hasMatch(name) ||
        _cellLikeName.hasMatch(name)) {
      throw ArgumentError.value(name, 'name',
          'must start with a letter or underscore, contain no spaces and not look like a cell reference');
    }
    final lower = name.toLowerCase();
    for (final sheet in _excel._sheetMap.values) {
      if (sheet._tables.any((t) => t.name.toLowerCase() == lower)) {
        throw ArgumentError.value(name, 'name', 'is already used by another table');
      }
    }
  }

  List<TableColumn> _columnsFromHeader(_CellRect rect, bool headerRow) {
    final used = <String>{};
    final result = <TableColumn>[];
    for (var c = rect.left; c <= rect.right; c++) {
      var name = '';
      if (headerRow) {
        final cell = _sheetData[rect.top]?[c];
        if (cell?.value != null) name = _headerText(cell!).trim();
      }
      if (name.isEmpty) name = 'Column${c - rect.left + 1}';
      var unique = name;
      var n = 2;
      while (!used.add(unique.toLowerCase())) {
        unique = '$name${n++}';
      }
      result.add(TableColumn(unique));
    }
    return result;
  }

  /// Writes the column names in the header row and the labels/formulas in
  /// the totals row, as Excel requires them to match the table definition.
  void _syncTableCells(ExcelTable table) {
    final rect = table._rect;
    for (var i = 0; i < table.columns.length; i++) {
      final column = table.columns[i];
      final col = rect.left + i;
      if (table.showHeaderRow) {
        final index = CellIndex.indexByColumnRow(columnIndex: col, rowIndex: rect.top);
        final current = _sheetData[rect.top]?[col]?.value;
        if (current is! TextCellValue || current.toString() != column.name) {
          updateCell(index, TextCellValue(column.name));
        }
      }
      if (table.showTotalsRow) {
        final value = _totalsCellValue(table, column);
        if (value != null) {
          updateCell(
              CellIndex.indexByColumnRow(columnIndex: col, rowIndex: rect.bottom), value);
        }
      }
    }
  }

  /// The label or formula the totals row shows for [column], if any.
  CellValue? _totalsCellValue(ExcelTable table, TableColumn column) {
    final code = column.totalsFunction.subtotalCode;
    if (column.totalsLabel != null) return TextCellValue(column.totalsLabel!);
    if (code != null) {
      return FormulaCellValue('SUBTOTAL($code,${table.columnReference(column.name)})');
    }
    if (column.totalsFunction == TableTotalsFunction.custom && column.totalsRowFormula != null) {
      return FormulaCellValue(column.totalsRowFormula!);
    }
    return null;
  }

  /// Throws when [table]'s totals row holds data it would overwrite or
  /// leave mixed with the totals; values the totals row writes anyway are
  /// allowed, so a removed table can be added again.
  void _checkTotalsRowIsFree(ExcelTable table) {
    final row = table.totalsRowIndex;
    if (row == null) return;
    final left = table._rect.left;
    for (var i = 0; i < table.columns.length; i++) {
      final value = _sheetData[row]?[left + i]?.value;
      if (value != null && value != _totalsCellValue(table, table.columns[i])) {
        throw ArgumentError.value(table.ref, 'range',
            'its last row (${row + 1}) has data the totals row would overwrite; '
            'leave an empty row at the end of the range for the totals');
      }
    }
  }

  /// Moves/resizes tables after inserting (`delta: 1`) or removing
  /// (`delta: -1`) the row or column at [index].
  void _shiftTables({bool rows = true, required int index, required int delta}) {
    if (_tables.isEmpty) return;
    final updated = <ExcelTable>[];
    for (final table in _tables) {
      final rect = table._rect;
      if (rows) {
        // Removing the header row drops the table; removing the totals row
        // turns the totals off.
        if (delta < 0 && table.showHeaderRow && index == rect.top) continue;
        var t = table;
        if (delta < 0 && table.showTotalsRow && index == rect.bottom) {
          t = t.copyWith(showTotalsRow: false);
        }
        final shifted = rect.shifted(rows: true, index: index, delta: delta);
        if (shifted == null) continue;
        updated.add(t.copyWith(ref: shifted.ref));
      } else {
        final shifted = rect.shifted(rows: false, index: index, delta: delta);
        if (shifted == null) continue;
        var columns = table.columns;
        final inside = delta > 0
            ? index > rect.left && index <= rect.right
            : index >= rect.left && index <= rect.right;
        if (inside) {
          columns = List.of(columns);
          final position = index - rect.left;
          if (delta > 0) {
            var n = columns.length + 1;
            final names = columns.map((c) => c.name.toLowerCase()).toSet();
            while (names.contains('column$n')) {
              n++;
            }
            columns.insert(position, TableColumn('Column$n'));
          } else {
            columns.removeAt(position);
          }
        }
        if (columns.isEmpty) continue;
        updated.add(table.copyWith(ref: shifted.ref, columns: columns));
      }
    }
    _tables
      ..clear()
      ..addAll(updated);
  }
}
