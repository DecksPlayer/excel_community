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
  static const black = BaseColors.black;
  static const white = BaseColors.white;

  // Base Colors
  static const black12 = BaseColors.black12;
  static const black26 = BaseColors.black26;
  static const black38 = BaseColors.black38;
  static const black45 = BaseColors.black45;
  static const black54 = BaseColors.black54;
  static const black87 = BaseColors.black87;

  static const white10 = BaseColors.white10;
  static const white12 = BaseColors.white12;
  static const white24 = BaseColors.white24;
  static const white30 = BaseColors.white30;
  static const white38 = BaseColors.white38;
  static const white54 = BaseColors.white54;
  static const white60 = BaseColors.white60;
  static const white70 = BaseColors.white70;

  static const grey = BaseColors.grey;
  static const grey50 = BaseColors.grey50;
  static const grey100 = BaseColors.grey100;
  static const grey200 = BaseColors.grey200;
  static const grey300 = BaseColors.grey300;
  static const grey350 = BaseColors.grey350;
  static const grey400 = BaseColors.grey400;
  static const grey600 = BaseColors.grey600;
  static const grey700 = BaseColors.grey700;
  static const grey800 = BaseColors.grey800;
  static const grey850 = BaseColors.grey850;
  static const grey900 = BaseColors.grey900;

  // Red Colors
  static const red = RedColors.red;
  static const red50 = RedColors.red50;
  static const red100 = RedColors.red100;
  static const red200 = RedColors.red200;
  static const red300 = RedColors.red300;
  static const red400 = RedColors.red400;
  static const red600 = RedColors.red600;
  static const red700 = RedColors.red700;
  static const red800 = RedColors.red800;
  static const red900 = RedColors.red900;

  static const pink = RedColors.pink;
  static const pink50 = RedColors.pink50;
  static const pink100 = RedColors.pink100;
  static const pink200 = RedColors.pink200;
  static const pink300 = RedColors.pink300;
  static const pink400 = RedColors.pink400;
  static const pink600 = RedColors.pink600;
  static const pink700 = RedColors.pink700;
  static const pink800 = RedColors.pink800;
  static const pink900 = RedColors.pink900;

  // Blue Colors
  static const blue = BlueColors.blue;
  static const blue50 = BlueColors.blue50;
  static const blue100 = BlueColors.blue100;
  static const blue200 = BlueColors.blue200;
  static const blue300 = BlueColors.blue300;
  static const blue400 = BlueColors.blue400;
  static const blue600 = BlueColors.blue600;
  static const blue700 = BlueColors.blue700;
  static const blue800 = BlueColors.blue800;
  static const blue900 = BlueColors.blue900;

  static const lightBlue = BlueColors.lightBlue;
  static const lightBlue50 = BlueColors.lightBlue50;
  static const lightBlue100 = BlueColors.lightBlue100;
  static const lightBlue200 = BlueColors.lightBlue200;
  static const lightBlue300 = BlueColors.lightBlue300;
  static const lightBlue400 = BlueColors.lightBlue400;
  static const lightBlue600 = BlueColors.lightBlue600;
  static const lightBlue700 = BlueColors.lightBlue700;
  static const lightBlue800 = BlueColors.lightBlue800;
  static const lightBlue900 = BlueColors.lightBlue900;

  // Green Colors
  static const green = GreenColors.green;
  static const green50 = GreenColors.green50;
  static const green100 = GreenColors.green100;
  static const green200 = GreenColors.green200;
  static const green300 = GreenColors.green300;
  static const green400 = GreenColors.green400;
  static const green600 = GreenColors.green600;
  static const green700 = GreenColors.green700;
  static const green800 = GreenColors.green800;
  static const green900 = GreenColors.green900;

  static const lightGreen = GreenColors.lightGreen;
  static const lightGreen50 = GreenColors.lightGreen50;
  static const lightGreen100 = GreenColors.lightGreen100;
  static const lightGreen200 = GreenColors.lightGreen200;
  static const lightGreen300 = GreenColors.lightGreen300;
  static const lightGreen400 = GreenColors.lightGreen400;
  static const lightGreen600 = GreenColors.lightGreen600;
  static const lightGreen700 = GreenColors.lightGreen700;
  static const lightGreen800 = GreenColors.lightGreen800;
  static const lightGreen900 = GreenColors.lightGreen900;

  // Yellow Orange Colors
  static const yellow = YellowOrangeColors.yellow;
  static const yellow50 = YellowOrangeColors.yellow50;
  static const yellow100 = YellowOrangeColors.yellow100;
  static const yellow200 = YellowOrangeColors.yellow200;
  static const yellow300 = YellowOrangeColors.yellow300;
  static const yellow400 = YellowOrangeColors.yellow400;
  static const yellow600 = YellowOrangeColors.yellow600;
  static const yellow700 = YellowOrangeColors.yellow700;
  static const yellow800 = YellowOrangeColors.yellow800;
  static const yellow900 = YellowOrangeColors.yellow900;

  static const amber = YellowOrangeColors.amber;
  static const amber50 = YellowOrangeColors.amber50;
  static const amber100 = YellowOrangeColors.amber100;
  static const amber200 = YellowOrangeColors.amber200;
  static const amber300 = YellowOrangeColors.amber300;
  static const amber400 = YellowOrangeColors.amber400;
  static const amber600 = YellowOrangeColors.amber600;
  static const amber700 = YellowOrangeColors.amber700;
  static const amber800 = YellowOrangeColors.amber800;
  static const amber900 = YellowOrangeColors.amber900;

  static const orange = YellowOrangeColors.orange;
  static const orange50 = YellowOrangeColors.orange50;
  static const orange100 = YellowOrangeColors.orange100;
  static const orange200 = YellowOrangeColors.orange200;
  static const orange300 = YellowOrangeColors.orange300;
  static const orange400 = YellowOrangeColors.orange400;
  static const orange600 = YellowOrangeColors.orange600;
  static const orange700 = YellowOrangeColors.orange700;
  static const orange800 = YellowOrangeColors.orange800;
  static const orange900 = YellowOrangeColors.orange900;

  static const deepOrange = YellowOrangeColors.deepOrange;
  static const deepOrange50 = YellowOrangeColors.deepOrange50;
  static const deepOrange100 = YellowOrangeColors.deepOrange100;
  static const deepOrange200 = YellowOrangeColors.deepOrange200;
  static const deepOrange300 = YellowOrangeColors.deepOrange300;
  static const deepOrange400 = YellowOrangeColors.deepOrange400;
  static const deepOrange600 = YellowOrangeColors.deepOrange600;
  static const deepOrange700 = YellowOrangeColors.deepOrange700;
  static const deepOrange800 = YellowOrangeColors.deepOrange800;
  static const deepOrange900 = YellowOrangeColors.deepOrange900;

  // Other Colors
  static const purple = OtherColors.purple;
  static const purple50 = OtherColors.purple50;
  static const purple100 = OtherColors.purple100;
  static const purple200 = OtherColors.purple200;
  static const purple300 = OtherColors.purple300;
  static const purple400 = OtherColors.purple400;
  static const purple600 = OtherColors.purple600;
  static const purple700 = OtherColors.purple700;
  static const purple800 = OtherColors.purple800;
  static const purple900 = OtherColors.purple900;

  static const deepPurple = OtherColors.deepPurple;
  static const deepPurple50 = OtherColors.deepPurple50;
  static const deepPurple100 = OtherColors.deepPurple100;
  static const deepPurple200 = OtherColors.deepPurple200;
  static const deepPurple300 = OtherColors.deepPurple300;
  static const deepPurple400 = OtherColors.deepPurple400;
  static const deepPurple600 = OtherColors.deepPurple600;
  static const deepPurple700 = OtherColors.deepPurple700;
  static const deepPurple800 = OtherColors.deepPurple800;
  static const deepPurple900 = OtherColors.deepPurple900;

  static const indigo = OtherColors.indigo;
  static const indigo50 = OtherColors.indigo50;
  static const indigo100 = OtherColors.indigo100;
  static const indigo200 = OtherColors.indigo200;
  static const indigo300 = OtherColors.indigo300;
  static const indigo400 = OtherColors.indigo400;
  static const indigo600 = OtherColors.indigo600;
  static const indigo700 = OtherColors.indigo700;
  static const indigo800 = OtherColors.indigo800;
  static const indigo900 = OtherColors.indigo900;

  static const cyan = OtherColors.cyan;
  static const cyan50 = OtherColors.cyan50;
  static const cyan100 = OtherColors.cyan100;
  static const cyan200 = OtherColors.cyan200;
  static const cyan300 = OtherColors.cyan300;
  static const cyan400 = OtherColors.cyan400;
  static const cyan600 = OtherColors.cyan600;
  static const cyan700 = OtherColors.cyan700;
  static const cyan800 = OtherColors.cyan800;
  static const cyan900 = OtherColors.cyan900;

  static const teal = OtherColors.teal;
  static const teal50 = OtherColors.teal50;
  static const teal100 = OtherColors.teal100;
  static const teal200 = OtherColors.teal200;
  static const teal300 = OtherColors.teal300;
  static const teal400 = OtherColors.teal400;
  static const teal600 = OtherColors.teal600;
  static const teal700 = OtherColors.teal700;
  static const teal800 = OtherColors.teal800;
  static const teal900 = OtherColors.teal900;

  static const lime = OtherColors.lime;
  static const lime50 = OtherColors.lime50;
  static const lime100 = OtherColors.lime100;
  static const lime200 = OtherColors.lime200;
  static const lime300 = OtherColors.lime300;
  static const lime400 = OtherColors.lime400;
  static const lime600 = OtherColors.lime600;
  static const lime700 = OtherColors.lime700;
  static const lime800 = OtherColors.lime800;
  static const lime900 = OtherColors.lime900;

  static const brown = OtherColors.brown;
  static const brown50 = OtherColors.brown50;
  static const brown100 = OtherColors.brown100;
  static const brown200 = OtherColors.brown200;
  static const brown300 = OtherColors.brown300;
  static const brown400 = OtherColors.brown400;
  static const brown600 = OtherColors.brown600;
  static const brown700 = OtherColors.brown700;
  static const brown800 = OtherColors.brown800;
  static const brown900 = OtherColors.brown900;

  static const blueGrey = OtherColors.blueGrey;
  static const blueGrey50 = OtherColors.blueGrey50;
  static const blueGrey100 = OtherColors.blueGrey100;
  static const blueGrey200 = OtherColors.blueGrey200;
  static const blueGrey300 = OtherColors.blueGrey300;
  static const blueGrey400 = OtherColors.blueGrey400;
  static const blueGrey600 = OtherColors.blueGrey600;
  static const blueGrey700 = OtherColors.blueGrey700;
  static const blueGrey800 = OtherColors.blueGrey800;
  static const blueGrey900 = OtherColors.blueGrey900;

  // Accent Colors
  static const redAccent = AccentColors.redAccent;
  static const redAccent100 = AccentColors.redAccent100;
  static const redAccent400 = AccentColors.redAccent400;
  static const redAccent700 = AccentColors.redAccent700;

  static const pinkAccent = AccentColors.pinkAccent;
  static const pinkAccent100 = AccentColors.pinkAccent100;
  static const pinkAccent400 = AccentColors.pinkAccent400;
  static const pinkAccent700 = AccentColors.pinkAccent700;

  static const purpleAccent = AccentColors.purpleAccent;
  static const purpleAccent100 = AccentColors.purpleAccent100;
  static const purpleAccent400 = AccentColors.purpleAccent400;
  static const purpleAccent700 = AccentColors.purpleAccent700;

  static const deepPurpleAccent = AccentColors.deepPurpleAccent;
  static const deepPurpleAccent100 = AccentColors.deepPurpleAccent100;
  static const deepPurpleAccent400 = AccentColors.deepPurpleAccent400;
  static const deepPurpleAccent700 = AccentColors.deepPurpleAccent700;

  static const indigoAccent = AccentColors.indigoAccent;
  static const indigoAccent100 = AccentColors.indigoAccent100;
  static const indigoAccent400 = AccentColors.indigoAccent400;
  static const indigoAccent700 = AccentColors.indigoAccent700;

  static const blueAccent = AccentColors.blueAccent;
  static const blueAccent100 = AccentColors.blueAccent100;
  static const blueAccent400 = AccentColors.blueAccent400;
  static const blueAccent700 = AccentColors.blueAccent700;

  static const lightBlueAccent = AccentColors.lightBlueAccent;
  static const lightBlueAccent100 = AccentColors.lightBlueAccent100;
  static const lightBlueAccent400 = AccentColors.lightBlueAccent400;
  static const lightBlueAccent700 = AccentColors.lightBlueAccent700;

  static const cyanAccent = AccentColors.cyanAccent;
  static const cyanAccent100 = AccentColors.cyanAccent100;
  static const cyanAccent400 = AccentColors.cyanAccent400;
  static const cyanAccent700 = AccentColors.cyanAccent700;

  static const tealAccent = AccentColors.tealAccent;
  static const tealAccent100 = AccentColors.tealAccent100;
  static const tealAccent400 = AccentColors.tealAccent400;
  static const tealAccent700 = AccentColors.tealAccent700;

  static const greenAccent = AccentColors.greenAccent;
  static const greenAccent100 = AccentColors.greenAccent100;
  static const greenAccent400 = AccentColors.greenAccent400;
  static const greenAccent700 = AccentColors.greenAccent700;

  static const lightGreenAccent = AccentColors.lightGreenAccent;
  static const lightGreenAccent100 = AccentColors.lightGreenAccent100;
  static const lightGreenAccent400 = AccentColors.lightGreenAccent400;
  static const lightGreenAccent700 = AccentColors.lightGreenAccent700;

  static const limeAccent = AccentColors.limeAccent;
  static const limeAccent100 = AccentColors.limeAccent100;
  static const limeAccent400 = AccentColors.limeAccent400;
  static const limeAccent700 = AccentColors.limeAccent700;

  static const yellowAccent = AccentColors.yellowAccent;
  static const yellowAccent100 = AccentColors.yellowAccent100;
  static const yellowAccent400 = AccentColors.yellowAccent400;
  static const yellowAccent700 = AccentColors.yellowAccent700;

  static const amberAccent = AccentColors.amberAccent;
  static const amberAccent100 = AccentColors.amberAccent100;
  static const amberAccent400 = AccentColors.amberAccent400;
  static const amberAccent700 = AccentColors.amberAccent700;

  static const orangeAccent = AccentColors.orangeAccent;
  static const orangeAccent100 = AccentColors.orangeAccent100;
  static const orangeAccent400 = AccentColors.orangeAccent400;
  static const orangeAccent700 = AccentColors.orangeAccent700;

  static const deepOrangeAccent = AccentColors.deepOrangeAccent;
  static const deepOrangeAccent100 = AccentColors.deepOrangeAccent100;
  static const deepOrangeAccent400 = AccentColors.deepOrangeAccent400;
  static const deepOrangeAccent700 = AccentColors.deepOrangeAccent700;

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
