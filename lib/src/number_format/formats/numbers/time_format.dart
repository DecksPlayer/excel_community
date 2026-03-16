part of excel_community;

sealed class TimeNumFormat extends NumFormat {
  const TimeNumFormat({
    required super.formatCode,
  });

  @override
  CellValue read(String v) {
    if (v == '0') {
      return const TimeCellValue(
        hour: 0,
        minute: 0,
        second: 0,
        millisecond: 0,
        microsecond: 0,
      );
    }
    var value = num.parse(v);
    if (value < 1) {
      var delta = value * 24 * 3600 * 1000;
      final time = Duration(milliseconds: delta.round());
      final date = DateTime.utc(0).add(time);
      return TimeCellValue(
        hour: date.hour,
        minute: date.minute,
        second: date.second,
        millisecond: date.millisecond,
        microsecond: date.microsecond,
      );
    }
    var delta = value * 24 * 3600 * 1000;
    var dateOffset = DateTime.utc(1899, 12, 30);
    final utcDate = dateOffset.add(Duration(milliseconds: delta.round()));
    if (!v.contains('.') || v.endsWith('.0')) {
      return DateCellValue(
        year: utcDate.year,
        month: utcDate.month,
        day: utcDate.day,
      );
    } else {
      return DateTimeCellValue(
        year: utcDate.year,
        month: utcDate.month,
        day: utcDate.day,
        hour: utcDate.hour,
        minute: utcDate.minute,
        second: utcDate.second,
        millisecond: utcDate.millisecond,
        microsecond: utcDate.microsecond,
      );
    }
  }

  String writeTime(TimeCellValue value) {
    final fractionOfDay =
        value.asDuration().inMilliseconds.toDouble() / (1000 * 3600 * 24);
    return fractionOfDay.toString();
  }

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => false,
        TextCellValue() => false,
        BoolCellValue() => false,
        DoubleCellValue() => false,
        DateCellValue() => false,
        DateTimeCellValue() => false,
        TimeCellValue() => true,
      };
}

class StandardTimeNumFormat extends TimeNumFormat implements StandardNumFormat {
  @override
  final int numFmtId;

  const StandardTimeNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  String toString() {
    return 'StandardTimeNumFormat($numFmtId, "$formatCode")';
  }
}

class CustomTimeNumFormat extends TimeNumFormat implements CustomNumFormat {
  const CustomTimeNumFormat({
    required super.formatCode,
  });

  @override
  String toString() {
    return 'CustomTimeNumFormat("$formatCode")';
  }
}
