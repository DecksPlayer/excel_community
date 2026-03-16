part of excel_community;

sealed class DateTimeNumFormat extends NumFormat {
  const DateTimeNumFormat({
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
    final value = num.parse(v);
    if (value < 1) {
      return TimeCellValue.fromFractionOfDay(value);
    }
    var delta = value * 24 * 3600 * 1000;
    var dateOffset = DateTime.utc(1899, 12, 30);
    final utcDate = dateOffset.add(Duration(milliseconds: delta.round()));
    if (!v.contains('.') || v.endsWith('.0')) {
      return DateCellValue.fromDateTime(utcDate);
    } else {
      return DateTimeCellValue.fromDateTime(utcDate);
    }
  }

  String writeDate(DateCellValue value) {
    var dateOffset = DateTime.utc(1899, 12, 30);
    final delta = value.asDateTimeUtc().difference(dateOffset);
    final dayFractions = delta.inMilliseconds.toDouble() / (1000 * 3600 * 24);
    return dayFractions.toString();
  }

  String writeDateTime(DateTimeCellValue value) {
    var dateOffset = DateTime.utc(1899, 12, 30);
    final delta = value.asDateTimeUtc().difference(dateOffset);
    final dayFractions = delta.inMilliseconds.toDouble() / (1000 * 3600 * 24);
    return dayFractions.toString();
  }

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => false,
        TextCellValue() => false,
        BoolCellValue() => false,
        DoubleCellValue() => false,
        DateCellValue() => true,
        DateTimeCellValue() => true,
        TimeCellValue() => false,
      };
}

class StandardDateTimeNumFormat extends DateTimeNumFormat
    implements StandardNumFormat {
  @override
  final int numFmtId;

  const StandardDateTimeNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  String toString() {
    return 'StandardDateTimeNumFormat($numFmtId, "$formatCode")';
  }
}

class CustomDateTimeNumFormat extends DateTimeNumFormat
    implements CustomNumFormat {
  const CustomDateTimeNumFormat({
    required super.formatCode,
  });

  @override
  String toString() {
    return 'CustomDateTimeNumFormat("$formatCode")';
  }
}
