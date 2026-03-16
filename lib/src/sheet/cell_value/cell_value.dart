part of excel_community;

/// Base class for all cell value types in an Excel sheet.
sealed class CellValue {
  const CellValue();

  /// The default format for this cell value type.
  NumFormat get defaultFormat;

  /// Writes the value as a string using the given [format].
  /// If [format] is null or incompatible, [defaultFormat] should be used.
  String write(NumFormat? format);
}
