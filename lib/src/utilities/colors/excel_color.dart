part of excel_community;

String _decimalToHexadecimal(int decimalVal) {
  if (decimalVal == 0) {
    return '0';
  }
  bool negative = false;
  if (decimalVal < 0) {
    negative = true;
    decimalVal *= -1;
  }
  String hexString = '';
  while (decimalVal > 0) {
    String hexVal = '';
    final int remainder = decimalVal % 16;
    decimalVal = decimalVal ~/ 16;
    if (_hexTable.containsKey(remainder)) {
      hexVal = _hexTable[remainder]!;
    } else {
      hexVal = remainder.toString();
    }
    hexString = hexVal + hexString;
  }
  return negative ? '-$hexString' : hexString;
}

bool _assertHexString(String hexString) {
  hexString = hexString.replaceAll('#', '').trim().toUpperCase();

  final bool isNegative = hexString[0] == '-';
  if (isNegative) hexString = hexString.substring(1);

  for (int i = 0; i < hexString.length; i++) {
    if (int.tryParse(hexString[i]) == null &&
        _hexTableReverse.containsKey(hexString[i]) == false) {
      return false;
    }
  }
  return true;
}

int _hexadecimalToDecimal(String hexString) {
  hexString = hexString.replaceAll('#', '').trim().toUpperCase();

  final bool isNegative = hexString[0] == '-';
  if (isNegative) hexString = hexString.substring(1);

  int decimalVal = 0;
  for (int i = 0; i < hexString.length; i++) {
    if (int.tryParse(hexString[i]) == null &&
        _hexTableReverse.containsKey(hexString[i]) == false) {
      throw Exception('Non-hex value was passed to the function');
    } else {
      decimalVal += (pow(16, hexString.length - i - 1) *
              (int.tryParse(hexString[i]) != null
                  ? int.parse(hexString[i])
                  : _hexTableReverse[hexString[i]]!))
          .toInt();
    }
  }
  return isNegative ? -1 * decimalVal : decimalVal;
}

const _hexTable = {
  10: 'A',
  11: 'B',
  12: 'C',
  13: 'D',
  14: 'E',
  15: 'F',
};

final _hexTableReverse = _hexTable.map((k, v) => MapEntry(v, k));

extension StringExt on String {
  /// Return [ExcelColor.black] if not a color hexadecimal
  ExcelColor get excelColor => this == 'none'
      ? ExcelColor.none
      : _assertHexString(this)
          ? ExcelColor.valuesAsMap[this] ?? ExcelColor._(this)
          : ExcelColor.black;
}

/// Copying from Flutter Material Color
class ExcelColor extends Equatable {
  const ExcelColor._(this._color, [this._name, this._type]);

  final String _color;
  final String? _name;
  final ColorType? _type;

  /// Return 'none' if [_color] is null, [black] if not match for safety
  String get colorHex =>
      _assertHexString(_color) || _color == 'none' ? _color : black.colorHex;

  /// Returns 6-character hex string (RRGGBB) for XML compatibility (mostly charts)
  String get colorHex6 {
    final hex = colorHex;
    if (hex == 'none') return 'none';
    if (hex.length >= 6) {
      return hex.substring(hex.length - 6);
    }
    return hex.padLeft(6, '0');
  }

  /// Return [black] if [_color] is not match for safety
  int get colorInt =>
      _assertHexString(_color) ? _hexadecimalToDecimal(_color) : black.colorInt;

  ColorType? get type => _type;

  String? get name => _name;

  /// Warning! Highly unsafe method.
  /// Can break your excel file if you do not know what you are doing
  factory ExcelColor.fromInt(int colorIntValue) =>
      ExcelColor._(_decimalToHexadecimal(colorIntValue));

  /// Warning! Highly unsafe method.
  /// Can break your excel file if you do not know what you are doing
  factory ExcelColor.fromHexString(String colorHexValue) =>
      ExcelColor._(colorHexValue);

  static const none = ExcelColor._('none');

  // Common colors kept in class for convenience
  static const black = ExcelColor._('FF000000', 'black', ColorType.color);
  static const white = ExcelColor._('FFFFFFFF', 'white', ColorType.color);

  // Constants mapping
  static List<ExcelColor> get values => [
        ...BaseColors.values,
        ...RedColors.values,
        ...BlueColors.values,
        ...GreenColors.values,
        ...YellowOrangeColors.values,
        ...OtherColors.values,
        ...AccentColors.values,
      ];

  static Map<String, ExcelColor> get valuesAsMap =>
      values.asMap().map((_, v) => MapEntry(v.colorHex, v));

  @override
  List<Object?> get props => [
        _name,
        _color,
        _type,
        colorHex,
        colorInt,
      ];
}

enum ColorType {
  color,
  material,
  materialAccent,
  ;
}
