part of excel_community;

sealed class NumericNumFormat extends NumFormat {
  const NumericNumFormat({
    required super.formatCode,
  });

  @override
  CellValue read(String v) {
    // check if scientific notation e.g. 1E-3
    final eIdx = v.indexOf('E');
    final decimalSeparatorIdx = v.indexOf('.');

    if (decimalSeparatorIdx == -1 && eIdx == -1) {
      return IntCellValue(int.parse(v));
    }

    // also read .0 (or even .00) as an int
    bool noActualDecimalPlaces = true;
    for (var idx = decimalSeparatorIdx + 1; idx < v.length; ++idx) {
      if (v[idx] != '0') {
        noActualDecimalPlaces = false;
        break;
      }
    }
    if (noActualDecimalPlaces) {
      return IntCellValue(int.parse(v.substring(0, decimalSeparatorIdx)));
    }

    return DoubleCellValue(double.parse(v));
  }

  String writeDouble(DoubleCellValue value) {
    return value.value.toString();
  }

  String writeInt(IntCellValue value) {
    return value.value.toString();
  }
}

class StandardNumericNumFormat extends NumericNumFormat
    implements StandardNumFormat {
  @override
  final int numFmtId;

  const StandardNumericNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => true,
        TextCellValue() => numFmtId == 0,
        BoolCellValue() => true,
        DoubleCellValue() => true,
        DateCellValue() => false,
        TimeCellValue() => false,
        DateTimeCellValue() => false,
      };

  @override
  String toString() {
    return 'StandardNumericNumFormat($numFmtId, "$formatCode")';
  }
}

class CustomNumericNumFormat extends NumericNumFormat
    implements CustomNumFormat {
  const CustomNumericNumFormat({
    required super.formatCode,
  });

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => true,
        TextCellValue() => false,
        BoolCellValue() => true,
        DoubleCellValue() => true,
        DateCellValue() => false,
        TimeCellValue() => false,
        DateTimeCellValue() => false,
      };

  @override
  String toString() {
    return 'CustomNumericNumFormat("$formatCode")';
  }
}
