part of excel_community;

/// Represents a boolean value in a cell.
class BoolCellValue extends CellValue {
  /// The boolean value.
  final bool value;

  /// Creates a new BoolCellValue.
  const BoolCellValue(this.value);

  @override
  NumFormat get defaultFormat => NumFormat.defaultBool;

  @override
  String write(NumFormat? format) {
    return value ? '1' : '0';
  }

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is BoolCellValue && other.value == value;
  }
}
