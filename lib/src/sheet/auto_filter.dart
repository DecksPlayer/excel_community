part of '../../excel_community.dart';

/// Comparison operators for custom auto-filter rules ([CustomFilterRule]).
enum FilterOperator {
  equal('equal'),
  lessThan('lessThan'),
  lessThanOrEqual('lessThanOrEqual'),
  notEqual('notEqual'),
  greaterThanOrEqual('greaterThanOrEqual'),
  greaterThan('greaterThan');

  final String value;
  const FilterOperator(this.value);

  static FilterOperator fromValue(String value) {
    return FilterOperator.values.firstWhere(
      (e) => e.value.toLowerCase() == value.toLowerCase(),
      orElse: () => FilterOperator.equal,
    );
  }
}

/// Represents a custom filter criterion in OpenXML `<customFilter>`.
class CustomFilterRule extends Equatable {
  final FilterOperator operator;
  final String val;

  const CustomFilterRule({
    this.operator = FilterOperator.equal,
    required this.val,
  });

  String toXmlString() {
    return '<customFilter operator="${operator.value}" val="${_escapeXml(val)}"/>';
  }

  @override
  List<Object?> get props => [operator, val];
}

/// Represents filter criteria on a single column in OpenXML `<filterColumn>`.
///
/// In OpenXML ECMA-376, [colId] is a zero-based column offset relative to the
/// first column of the AutoFilter range (0 refers to the start column).
class FilterColumn extends Equatable {
  /// Zero-based column index relative to the start of the AutoFilter range.
  final int colId;

  /// Whether the filter dropdown button in Excel is hidden for this column.
  final bool? hiddenButton;

  /// Whether the filter dropdown button is shown for this column.
  final bool? showButton;

  /// List of matching values for standard filters (`<filters><filter val="..."/></filters>`).
  final List<String> filterValues;

  /// Whether blank cells are matched in `<filters blank="1">`.
  final bool blank;

  /// Custom filter rules (`<customFilters>`).
  final List<CustomFilterRule> customFilters;

  /// When multiple custom filters are present, whether they are joined by AND (`true`) or OR (`false`).
  final bool customFiltersAnd;

  /// Preserved raw inner XML for complex filter column children
  /// (e.g. `<top10>`, `<colorFilter>`, `<dynamicFilter>`, `<iconFilter>`, `<dateGroupItem>`).
  final String? customXml;

  const FilterColumn({
    required this.colId,
    this.hiddenButton,
    this.showButton,
    this.filterValues = const [],
    this.blank = false,
    this.customFilters = const [],
    this.customFiltersAnd = false,
    this.customXml,
  });

  FilterColumn copyWith({
    int? colId,
    bool? hiddenButton,
    bool? showButton,
    List<String>? filterValues,
    bool? blank,
    List<CustomFilterRule>? customFilters,
    bool? customFiltersAnd,
    String? customXml,
  }) {
    return FilterColumn(
      colId: colId ?? this.colId,
      hiddenButton: hiddenButton ?? this.hiddenButton,
      showButton: showButton ?? this.showButton,
      filterValues: filterValues ?? this.filterValues,
      blank: blank ?? this.blank,
      customFilters: customFilters ?? this.customFilters,
      customFiltersAnd: customFiltersAnd ?? this.customFiltersAnd,
      customXml: customXml ?? this.customXml,
    );
  }

  String toXmlString() {
    if (customXml != null && customXml!.trim().isNotEmpty) {
      final sb = StringBuffer('<filterColumn colId="$colId"');
      if (hiddenButton != null) {
        sb.write(' hiddenButton="${hiddenButton! ? '1' : '0'}"');
      }
      if (showButton != null) {
        sb.write(' showButton="${showButton! ? '1' : '0'}"');
      }
      sb.write('>$customXml</filterColumn>');
      return sb.toString();
    }

    final hasFilters = filterValues.isNotEmpty || blank;
    final hasCustom = customFilters.isNotEmpty;

    if (!hasFilters && !hasCustom) {
      final sb = StringBuffer('<filterColumn colId="$colId"');
      if (hiddenButton != null) {
        sb.write(' hiddenButton="${hiddenButton! ? '1' : '0'}"');
      }
      if (showButton != null) {
        sb.write(' showButton="${showButton! ? '1' : '0'}"');
      }
      sb.write('/>');
      return sb.toString();
    }

    final sb = StringBuffer('<filterColumn colId="$colId"');
    if (hiddenButton != null) {
      sb.write(' hiddenButton="${hiddenButton! ? '1' : '0'}"');
    }
    if (showButton != null) {
      sb.write(' showButton="${showButton! ? '1' : '0'}"');
    }
    sb.write('>');

    if (hasFilters) {
      if (filterValues.isEmpty) {
        sb.write('<filters');
        if (blank) {
          sb.write(' blank="1"');
        }
        sb.write('/>');
      } else {
        sb.write('<filters');
        if (blank) {
          sb.write(' blank="1"');
        }
        sb.write('>');
        for (final val in filterValues) {
          sb.write('<filter val="${_escapeXml(val)}"/>');
        }
        sb.write('</filters>');
      }
    }

    if (hasCustom) {
      sb.write('<customFilters');
      if (customFiltersAnd) {
        sb.write(' and="1"');
      }
      sb.write('>');
      for (final rule in customFilters) {
        sb.write(rule.toXmlString());
      }
      sb.write('</customFilters>');
    }

    sb.write('</filterColumn>');
    return sb.toString();
  }

  @override
  List<Object?> get props => [
        colId,
        hiddenButton,
        showButton,
        filterValues,
        blank,
        customFilters,
        customFiltersAnd,
        customXml,
      ];
}

/// Represents the OpenXML `<autoFilter>` element for a worksheet.
///
/// Contains the cell range reference ([ref]) and optional column filter definitions ([filterColumns]).
class AutoFilter extends Equatable {
  /// The cell range reference string, e.g. `"A1:D10"`.
  final String ref;

  /// Optional column filters applied within this AutoFilter range.
  final List<FilterColumn> filterColumns;

  /// Preserved raw inner XML from an existing spreadsheet.
  final String? customXml;

  AutoFilter({
    required String ref,
    List<FilterColumn>? filterColumns,
    this.customXml,
  })  : ref = ref.trim().toUpperCase(),
        filterColumns = filterColumns != null
            ? List<FilterColumn>.unmodifiable(filterColumns)
            : const [];

  /// Creates an [AutoFilter] spanning between [start] and [end] cell indices.
  ///
  /// Coordinates are automatically normalized so [start] is upper-left and [end] is lower-right.
  factory AutoFilter.fromRange({
    required CellIndex start,
    required CellIndex end,
    List<FilterColumn>? filterColumns,
    String? customXml,
  }) {
    final minCol = min(start.columnIndex, end.columnIndex);
    final maxCol = max(start.columnIndex, end.columnIndex);
    final minRow = min(start.rowIndex, end.rowIndex);
    final maxRow = max(start.rowIndex, end.rowIndex);
    final rangeStr = '${getCellId(minCol, minRow)}:${getCellId(maxCol, maxRow)}';
    return AutoFilter(
      ref: rangeStr,
      filterColumns: filterColumns,
      customXml: customXml,
    );
  }

  /// Creates an [AutoFilter] from a cell range string (e.g. `"A1:D10"`).
  factory AutoFilter.fromRangeString(
    String ref, {
    List<FilterColumn>? filterColumns,
    String? customXml,
  }) {
    return AutoFilter(
      ref: ref,
      filterColumns: filterColumns,
      customXml: customXml,
    );
  }

  /// Returns the upper-left [CellIndex] of this auto filter.
  CellIndex get startCell {
    final parts = ref.split(':');
    return CellIndex.indexByString(parts.first);
  }

  /// Returns the lower-right [CellIndex] of this auto filter.
  CellIndex get endCell {
    final parts = ref.split(':');
    return parts.length > 1
        ? CellIndex.indexByString(parts.last)
        : CellIndex.indexByString(parts.first);
  }

  /// The 0-based column index of the upper-left cell.
  int get startColumn => startCell.columnIndex;

  /// The 0-based row index of the upper-left cell.
  int get startRow => startCell.rowIndex;

  /// The 0-based column index of the lower-right cell.
  int get endColumn => endCell.columnIndex;

  /// The 0-based row index of the lower-right cell.
  int get endRow => endCell.rowIndex;

  /// Number of columns included in this AutoFilter.
  int get columnCount => (endColumn - startColumn).abs() + 1;

  /// Number of rows included in this AutoFilter.
  int get rowCount => (endRow - startRow).abs() + 1;

  /// Returns whether the given [cell] is within this AutoFilter range.
  bool containsCell(CellIndex cell) {
    final minCol = min(startColumn, endColumn);
    final maxCol = max(startColumn, endColumn);
    final minR = min(startRow, endRow);
    final maxR = max(startRow, endRow);
    return cell.columnIndex >= minCol &&
        cell.columnIndex <= maxCol &&
        cell.rowIndex >= minR &&
        cell.rowIndex <= maxR;
  }

  /// Returns whether the given cell ID (e.g. `"B2"`) is within this AutoFilter range.
  bool containsCellId(String cellId) =>
      containsCell(CellIndex.indexByString(cellId));

  AutoFilter copyWith({
    String? ref,
    List<FilterColumn>? filterColumns,
    String? customXml,
  }) {
    return AutoFilter(
      ref: ref ?? this.ref,
      filterColumns: filterColumns ?? this.filterColumns,
      customXml: customXml ?? this.customXml,
    );
  }

  /// Serializes this [AutoFilter] to its OpenXML `<autoFilter>` representation.
  String toXmlString() {
    if (filterColumns.isNotEmpty) {
      final sb = StringBuffer('<autoFilter ref="$ref">');
      for (final col in filterColumns) {
        sb.write(col.toXmlString());
      }
      sb.write('</autoFilter>');
      return sb.toString();
    } else if (customXml != null && customXml!.trim().isNotEmpty) {
      return '<autoFilter ref="$ref">$customXml</autoFilter>';
    } else {
      return '<autoFilter ref="$ref"/>';
    }
  }

  @override
  List<Object?> get props => [ref, filterColumns, customXml];
}
