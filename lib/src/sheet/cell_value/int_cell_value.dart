part of excel_community;

/// Represents an integer value in a cell.
class IntCellValue extends CellValue {
  /// The integer value.
  final int value;

  /// Creates a new IntCellValue.
  const IntCellValue(this.value);

  @override
  NumFormat get defaultFormat => NumFormat.defaultNumeric;

  @override
  String write(NumFormat? format) {
    if (format is NumericNumFormat) {
      return format.writeInt(this);
    }
    return (defaultFormat as NumericNumFormat).writeInt(this);
  }

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is IntCellValue && other.value == value;
  }
}
