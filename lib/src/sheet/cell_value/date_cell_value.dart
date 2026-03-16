part of excel_community;

/// Represents a date value in a cell.
class DateCellValue extends CellValue {
  /// Year component.
  final int year;
  /// Month component (1-12).
  final int month;
  /// Day component (1-31).
  final int day;

  /// Creates a new DateCellValue.
  const DateCellValue({
    required this.year,
    required this.month,
    required this.day,
  })  : assert(month <= 12 && month >= 1),
        assert(day <= 31 && day >= 1);

  /// Creates a DateCellValue from a [DateTime] object.
  DateCellValue.fromDateTime(DateTime dt)
      : year = dt.year,
        month = dt.month,
        day = dt.day;

  @override
  NumFormat get defaultFormat => NumFormat.defaultDate;

  @override
  String write(NumFormat? format) {
    if (format is DateTimeNumFormat) {
      return format.writeDate(this);
    }
    return (defaultFormat as DateTimeNumFormat).writeDate(this);
  }

  /// Returns the date as a local [DateTime] object.
  DateTime asDateTimeLocal() {
    return DateTime(year, month, day);
  }

  /// Returns the date as a UTC [DateTime] object.
  DateTime asDateTimeUtc() {
    return DateTime.utc(year, month, day);
  }

  @override
  String toString() {
    return asDateTimeUtc().toIso8601String();
  }

  @override
  int get hashCode => Object.hash(runtimeType, year, month, day);

  @override
  operator ==(Object other) {
    return other is DateCellValue &&
        other.year == year &&
        other.month == month &&
        other.day == day;
  }
}
