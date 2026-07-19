part of '../../excel_community.dart';

/// Supported aggregation functions for Pivot Table values.
enum PivotValueFunction {
  sum,
  count,
  average,
  max,
  min,
  product,
  countNums,
  stdDev,
  stdDevp,
  varVal,
  varp,
}

/// Represents a single value/data field in a Pivot Table.
class PivotTableValue {
  final String field;
  final PivotValueFunction function;
  final String? customName;

  PivotTableValue({
    required this.field,
    this.function = PivotValueFunction.sum,
    this.customName,
  });
}

/// Represents a Pivot Table configuration to be generated programmatically.
class PivotTable {
  final String name;
  final String sourceSheet;
  final String sourceRange; // e.g. "A1:B4"
  final CellIndex targetCell; // Top-left position on the target sheet
  final List<String> rows;
  final List<String> columns;
  final List<PivotTableValue> values;

  PivotTable({
    required this.name,
    required this.sourceSheet,
    required this.sourceRange,
    required this.targetCell,
    this.rows = const [],
    this.columns = const [],
    this.values = const [],
  });
}
