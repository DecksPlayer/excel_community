part of '../../../excel_community.dart';

/// Represents a formula in a cell.
class FormulaCellValue extends CellValue {
  /// The formula string (e.g., "=SUM(A1:A10)").
  final String formula;

  /// The last value Excel computed for this formula and stored in the XLSX's
  /// `<v>` element, if the source file had one.
  ///
  /// This is `null` for formulas created programmatically (before saving) and
  /// is never recalculated by this library — it only reflects whatever value
  /// was cached in the file at the time it was read.
  final CellValue? cachedValue;

  /// Creates a new FormulaCellValue.
  const FormulaCellValue(this.formula, {this.cachedValue});

  @override
  NumFormat get defaultFormat => NumFormat.standard_0;

  @override
  String write(NumFormat? format) {
    // Formulas usually don't rely on the format for their string representation in <v>
    // because they are often just placeholders or the evaluated result is handled separately.
    // In XLSX, <v> for formula cells is often empty and computed on open.
    return '';
  }

  @override
  String toString() {
    return formula;
  }

  @override
  int get hashCode => Object.hash(runtimeType, formula, cachedValue);

  @override
  operator ==(Object other) {
    return other is FormulaCellValue &&
        other.formula == formula &&
        other.cachedValue == cachedValue;
  }
}
