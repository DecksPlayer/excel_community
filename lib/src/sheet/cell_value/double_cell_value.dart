part of excel_community;

/// Represents a double (floating-point) value in a cell.
class DoubleCellValue extends CellValue {
  /// The double value.
  final double value;

  /// Creates a new DoubleCellValue.
  const DoubleCellValue(this.value);

  @override
  NumFormat get defaultFormat => NumFormat.defaultFloat;

  @override
  String write(NumFormat? format) {
    if (format is NumericNumFormat) {
      return format.writeDouble(this);
    }
    return (defaultFormat as NumericNumFormat).writeDouble(this);
  }

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is DoubleCellValue && other.value == value;
  }
}
