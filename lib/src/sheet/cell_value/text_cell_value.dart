part of excel_community;

/// Represents a text value in a cell, supporting rich text through spans.
class TextCellValue extends CellValue {
  /// The text content as a [TextSpan].
  final TextSpan value;

  /// Creates a new TextCellValue from a simple string.
  TextCellValue(String text) : value = TextSpan(text: text);

  /// Creates a new TextCellValue from a [TextSpan] for rich text support.
  TextCellValue.span(this.value);

  @override
  NumFormat get defaultFormat => NumFormat.standard_0;

  @override
  String write(NumFormat? format) {
    return value.toString();
  }

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is TextCellValue && other.value == value;
  }
}
