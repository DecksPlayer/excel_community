part of excel_community;

/// Excel does not know if this is UTC or not. Use methods [asDateTimeLocal]
/// or [asDateTimeUtc] to get the DateTime object you prefer.
class DateTimeCellValue extends CellValue {
  final int year;
  final int month;
  final int day;
  final int hour;
  final int minute;
  final int second;
  final int millisecond;
  final int microsecond;

  const DateTimeCellValue({
    required this.year,
    required this.month,
    required this.day,
    required this.hour,
    required this.minute,
    this.second = 0,
    this.millisecond = 0,
    this.microsecond = 0,
  })  : assert(month <= 12 && month >= 1),
        assert(day <= 31 && day >= 1),
        assert(hour <= 24 && hour >= 0),
        assert(minute <= 60 && minute >= 0),
        assert(second <= 60 && second >= 0),
        assert(millisecond <= 1000 && millisecond >= 0),
        assert(microsecond <= 1000 && microsecond >= 0);

  DateTimeCellValue.fromDateTime(DateTime date)
      : year = date.year,
        month = date.month,
        day = date.day,
        hour = date.hour,
        minute = date.minute,
        second = date.second,
        millisecond = date.millisecond,
        microsecond = date.microsecond;

  @override
  NumFormat get defaultFormat => NumFormat.defaultDateTime;

  @override
  String write(NumFormat? format) {
    if (format is DateTimeNumFormat) {
      return format.writeDateTime(this);
    }
    return (defaultFormat as DateTimeNumFormat).writeDateTime(this);
  }

  DateTime asDateTimeLocal() {
    return DateTime(
        year, month, day, hour, minute, second, millisecond, microsecond);
  }

  DateTime asDateTimeUtc() {
    return DateTime.utc(
        year, month, day, hour, minute, second, millisecond, microsecond);
  }

  @override
  String toString() {
    return asDateTimeUtc().toIso8601String();
  }

  @override
  int get hashCode => Object.hash(
        runtimeType,
        year,
        month,
        day,
        hour,
        minute,
        second,
        millisecond,
        microsecond,
      );

  @override
  operator ==(Object other) {
    return other is DateTimeCellValue &&
        other.year == year &&
        other.month == month &&
        other.day == day &&
        other.hour == hour &&
        other.minute == minute &&
        other.second == second &&
        other.millisecond == millisecond &&
        other.microsecond == microsecond;
  }
}
