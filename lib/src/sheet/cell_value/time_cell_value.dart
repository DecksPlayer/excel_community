part of excel_community;

class TimeCellValue extends CellValue {
  final int hour;
  final int minute;
  final int second;
  final int millisecond;
  final int microsecond;

  const TimeCellValue({
    this.hour = 0,
    this.minute = 0,
    this.second = 0,
    this.millisecond = 0,
    this.microsecond = 0,
  })  : assert(hour >= 0),
        assert(minute <= 60 && minute >= 0),
        assert(second <= 60 && second >= 0),
        assert(millisecond <= 1000 && millisecond >= 0),
        assert(microsecond <= 1000 && microsecond >= 0);

  /// [fractionOfDay]=1.0 is 24 hours, 0.5 is 12 hours and so on.
  factory TimeCellValue.fromFractionOfDay(num fractionOfDay) {
    var duration =
        Duration(milliseconds: (fractionOfDay * 24 * 3600 * 1000).round());
    return TimeCellValue.fromDuration(duration);
  }

  factory TimeCellValue.fromDuration(Duration duration) {
    final someUtcDate = DateTime.utc(0).add(duration);
    return TimeCellValue(
      hour: someUtcDate.hour,
      minute: someUtcDate.minute,
      second: someUtcDate.second,
      millisecond: someUtcDate.millisecond,
      microsecond: someUtcDate.microsecond,
    );
  }

  TimeCellValue.fromTimeOfDateTime(DateTime dt)
      : hour = dt.hour,
        minute = dt.minute,
        second = dt.second,
        millisecond = dt.millisecond,
        microsecond = dt.microsecond;

  Duration asDuration() {
    return Duration(
      hours: hour,
      minutes: minute,
      seconds: second,
      milliseconds: millisecond,
      microseconds: microsecond,
    );
  }

  @override
  NumFormat get defaultFormat => NumFormat.defaultTime;

  @override
  String write(NumFormat? format) {
    if (format is TimeNumFormat) {
      return format.writeTime(this);
    }
    return (defaultFormat as TimeNumFormat).writeTime(this);
  }

  @override
  String toString() {
    return '${_twoDigits(hour)}:${_twoDigits(minute)}:${_twoDigits(second)}';
  }

  @override
  int get hashCode => Object.hash(
        runtimeType,
        hour,
        minute,
        second,
        millisecond,
        microsecond,
      );

  @override
  operator ==(Object other) {
    return other is TimeCellValue &&
        other.hour == hour &&
        other.minute == minute &&
        other.second == second &&
        other.millisecond == millisecond &&
        other.microsecond == microsecond;
  }
}
