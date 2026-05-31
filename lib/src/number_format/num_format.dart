part of '../../excel_community.dart';

Map<V, K> _createInverseMap<K, V>(Map<K, V> map) {
  final inverse = <V, K>{};
  for (var entry in map.entries) {
    assert(!inverse.containsKey(entry.value), 'map values are not unique');
    inverse[entry.value] = entry.key;
  }
  return inverse;
}

class NumFormatMaintainer {
  static const int _firstCustomFmtId = 164;
  int _nextFmtId = _firstCustomFmtId;
  Map<int, NumFormat> _map = {..._standardNumFormats};
  Map<NumFormat, int> _inverseMap = _createInverseMap(_standardNumFormats);

  void add(int numFmtId, CustomNumFormat format) {
    // Allow overriding any ID. Excel files can explicitly declare built-in
    // format IDs (0–163) in numFmts to override them for a specific file.
    _map[numFmtId] = format;
    _inverseMap[format] = numFmtId;
    if (numFmtId >= _nextFmtId) {
      _nextFmtId = numFmtId + 1;
    }
  }

  int findOrAdd(CustomNumFormat format) {
    var fmtId = _inverseMap[format];
    if (fmtId != null) {
      return fmtId;
    }
    fmtId = _nextFmtId;
    _nextFmtId++;
    _map[fmtId] = format;
    return fmtId;
  }

  void clear() {
    _nextFmtId = _firstCustomFmtId;
    _map = {..._standardNumFormats};
    _inverseMap = _createInverseMap(_standardNumFormats);
  }

  NumFormat? getByNumFmtId(int numFmtId) {
    final result = _map[numFmtId];
    if (result != null) return result;
    // IDs 0–163 are built-in. If not in our map it is a locale-specific
    // format (50–163) that varies per Excel installation. Fall back to
    // General (standard_0) so the file opens without crashing.
    if (numFmtId >= 0 && numFmtId < 164) return NumFormat.standard_0;
    return null;
  }
}

sealed class NumFormat {
  final String formatCode;

  static const defaultNumeric = standard_1;
  static const defaultFloat = standard_2;
  static const defaultBool = standard_0;
  static const defaultDate = standard_14;
  static const defaultTime = standard_20;
  static const defaultDateTime = standard_22;

  static const standard_0 = StandardFormats.standard_0;
  static const standard_1 = StandardFormats.standard_1;
  static const standard_2 = StandardFormats.standard_2;
  static const standard_3 = StandardFormats.standard_3;
  static const standard_4 = StandardFormats.standard_4;
  // Currency (5–8)
  static const standard_5 = StandardFormats.standard_5;
  static const standard_6 = StandardFormats.standard_6;
  static const standard_7 = StandardFormats.standard_7;
  static const standard_8 = StandardFormats.standard_8;
  static const standard_9 = StandardFormats.standard_9;
  static const standard_10 = StandardFormats.standard_10;
  static const standard_11 = StandardFormats.standard_11;
  static const standard_12 = StandardFormats.standard_12;
  static const standard_13 = StandardFormats.standard_13;
  static const standard_14 = StandardFormats.standard_14;
  static const standard_15 = StandardFormats.standard_15;
  static const standard_16 = StandardFormats.standard_16;
  static const standard_17 = StandardFormats.standard_17;
  static const standard_18 = StandardFormats.standard_18;
  static const standard_19 = StandardFormats.standard_19;
  static const standard_20 = StandardFormats.standard_20;
  static const standard_21 = StandardFormats.standard_21;
  static const standard_22 = StandardFormats.standard_22;
  // Reserved (23–26)
  static const standard_23 = StandardFormats.standard_23;
  static const standard_24 = StandardFormats.standard_24;
  static const standard_25 = StandardFormats.standard_25;
  static const standard_26 = StandardFormats.standard_26;
  // CJK locale Date/Time (27–36)
  static const standard_27 = StandardFormats.standard_27;
  static const standard_28 = StandardFormats.standard_28;
  static const standard_29 = StandardFormats.standard_29;
  static const standard_30 = StandardFormats.standard_30;
  static const standard_31 = StandardFormats.standard_31;
  static const standard_32 = StandardFormats.standard_32;
  static const standard_33 = StandardFormats.standard_33;
  static const standard_34 = StandardFormats.standard_34;
  static const standard_35 = StandardFormats.standard_35;
  static const standard_36 = StandardFormats.standard_36;
  static const standard_37 = StandardFormats.standard_37;
  static const standard_38 = StandardFormats.standard_38;
  static const standard_39 = StandardFormats.standard_39;
  static const standard_40 = StandardFormats.standard_40;
  // Accounting with fill character (41–44)
  static const standard_41 = StandardFormats.standard_41;
  static const standard_42 = StandardFormats.standard_42;
  static const standard_43 = StandardFormats.standard_43;
  static const standard_44 = StandardFormats.standard_44;
  static const standard_45 = StandardFormats.standard_45;
  static const standard_46 = StandardFormats.standard_46;
  static const standard_47 = StandardFormats.standard_47;
  static const standard_48 = StandardFormats.standard_48;
  static const standard_49 = StandardFormats.standard_49;

  const NumFormat({
    required this.formatCode,
  });

  static CustomNumFormat custom({
    required String formatCode,
  }) {
    if (formatCode == 'General') {
      return CustomNumericNumFormat(formatCode: 'General');
    }

    if (_formatCodeLooksLikeDateTime(formatCode)) {
      return CustomDateTimeNumFormat(formatCode: formatCode);
    } else {
      return CustomNumericNumFormat(formatCode: formatCode);
    }
  }

  CellValue read(String v);

  @override
  int get hashCode => Object.hash(runtimeType, formatCode);

  @override
  operator ==(Object other) =>
      other.runtimeType == runtimeType &&
      (other as NumFormat).formatCode == formatCode;

  bool accepts(CellValue? value);

  static NumFormat defaultFor(CellValue? value) => switch (value) {
        null || FormulaCellValue() || TextCellValue() => NumFormat.standard_0,
        IntCellValue() => NumFormat.defaultNumeric,
        DoubleCellValue() => NumFormat.defaultFloat,
        DateCellValue() => NumFormat.defaultDate,
        BoolCellValue() => NumFormat.defaultBool,
        TimeCellValue() => NumFormat.defaultTime,
        DateTimeCellValue() => NumFormat.defaultDateTime,
      };
}

bool _formatCodeLooksLikeDateTime(String formatCode) {
  // for comparison, remove any character that is quoted or escaped
  var inEscape = false;
  var inQuotes = false;
  for (var i = 0; i < formatCode.length; ++i) {
    final c = formatCode[i];
    if (inEscape) {
      inEscape = false;
      continue;
    } else if (c == '\\') {
      inEscape = true;
      continue;
    }
    if (inQuotes) {
      if (c == '"') {
        inQuotes = false;
      }
      continue;
    } else if (c == '"') {
      inQuotes = true;
      continue;
    }

    switch (c) {
      case 'y':
      case 'm':
      case 'd':
      case 'h':
      case 's':
        return true;
      case ';':
        // separator only exists for decimal formats
        return false;
      default:
        break;
    }
  }
  return false;
}

sealed class StandardNumFormat implements NumFormat {
  int get numFmtId;
}

sealed class CustomNumFormat implements NumFormat {
  String get formatCode;
}
