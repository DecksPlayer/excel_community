part of excel_community;

class BaseColors {
  static const black = ExcelColor._('FF000000', 'black', ColorType.color);
  static const black12 = ExcelColor._('1F000000', 'black12', ColorType.color);
  static const black26 = ExcelColor._('42000000', 'black26', ColorType.color);
  static const black38 = ExcelColor._('61000000', 'black38', ColorType.color);
  static const black45 = ExcelColor._('73000000', 'black45', ColorType.color);
  static const black54 = ExcelColor._('8A000000', 'black54', ColorType.color);
  static const black87 = ExcelColor._('DD000000', 'black87', ColorType.color);
  
  static const white = ExcelColor._('FFFFFFFF', 'white', ColorType.color);
  static const white10 = ExcelColor._('1AFFFFFF', 'white10', ColorType.color);
  static const white12 = ExcelColor._('1FFFFFFF', 'white12', ColorType.color);
  static const white24 = ExcelColor._('3DFFFFFF', 'white24', ColorType.color);
  static const white30 = ExcelColor._('4DFFFFFF', 'white30', ColorType.color);
  static const white38 = ExcelColor._('62FFFFFF', 'white38', ColorType.color);
  static const white54 = ExcelColor._('8AFFFFFF', 'white54', ColorType.color);
  static const white60 = ExcelColor._('99FFFFFF', 'white60', ColorType.color);
  static const white70 = ExcelColor._('B3FFFFFF', 'white70', ColorType.color);

  static const grey = ExcelColor._('FF9E9E9E', 'grey', ColorType.material);
  static const grey50 = ExcelColor._('FFFAFAFA', 'grey50', ColorType.material);
  static const grey100 = ExcelColor._('FFF5F5F5', 'grey100', ColorType.material);
  static const grey200 = ExcelColor._('FFEEEEEE', 'grey200', ColorType.material);
  static const grey300 = ExcelColor._('FFE0E0E0', 'grey300', ColorType.material);
  static const grey350 = ExcelColor._('FFD6D6D6', 'grey350', ColorType.material);
  static const grey400 = ExcelColor._('FFBDBDBD', 'grey400', ColorType.material);
  static const grey600 = ExcelColor._('FF757575', 'grey600', ColorType.material);
  static const grey700 = ExcelColor._('FF616161', 'grey700', ColorType.material);
  static const grey800 = ExcelColor._('FF424242', 'grey800', ColorType.material);
  static const grey850 = ExcelColor._('FF303030', 'grey850', ColorType.material);
  static const grey900 = ExcelColor._('FF212121', 'grey900', ColorType.material);

  static List<ExcelColor> get values => [
        black,
        black12,
        black26,
        black38,
        black45,
        black54,
        black87,
        white,
        white10,
        white12,
        white24,
        white30,
        white38,
        white54,
        white60,
        white70,
        grey,
        grey50,
        grey100,
        grey200,
        grey300,
        grey350,
        grey400,
        grey600,
        grey700,
        grey800,
        grey850,
        grey900,
      ];
}
