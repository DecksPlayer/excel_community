part of excel_community;

const Map<int, NumFormat> _standardNumFormats = {
  0: NumFormat.standard_0,
  1: NumFormat.standard_1,
  2: NumFormat.standard_2,
  3: NumFormat.standard_3,
  4: NumFormat.standard_4,
  9: NumFormat.standard_9,
  10: NumFormat.standard_10,
  11: NumFormat.standard_11,
  12: NumFormat.standard_12,
  13: NumFormat.standard_13,
  14: NumFormat.standard_14,
  15: NumFormat.standard_15,
  16: NumFormat.standard_16,
  17: NumFormat.standard_17,
  18: NumFormat.standard_18,
  19: NumFormat.standard_19,
  20: NumFormat.standard_20,
  21: NumFormat.standard_21,
  22: NumFormat.standard_22,
  37: NumFormat.standard_37,
  38: NumFormat.standard_38,
  39: NumFormat.standard_39,
  40: NumFormat.standard_40,
  45: NumFormat.standard_45,
  46: NumFormat.standard_46,
  47: NumFormat.standard_47,
  48: NumFormat.standard_48,
  49: NumFormat.standard_49,
};

/// Helper to keep NumFormat clean
class StandardFormats {
  static const standard_0 =
      StandardNumericNumFormat._(numFmtId: 0, formatCode: 'General');
  static const standard_1 =
      StandardNumericNumFormat._(numFmtId: 1, formatCode: "0");
  static const standard_2 =
      StandardNumericNumFormat._(numFmtId: 2, formatCode: "0.00");
  static const standard_3 =
      StandardNumericNumFormat._(numFmtId: 3, formatCode: "#,##0");
  static const standard_4 =
      StandardNumericNumFormat._(numFmtId: 4, formatCode: "#,##0.00");
  static const standard_9 =
      StandardNumericNumFormat._(numFmtId: 9, formatCode: "0%");
  static const standard_10 =
      StandardNumericNumFormat._(numFmtId: 10, formatCode: "0.00%");
  static const standard_11 =
      StandardNumericNumFormat._(numFmtId: 11, formatCode: "0.00E+00");
  static const standard_12 =
      StandardNumericNumFormat._(numFmtId: 12, formatCode: "# ?/?");
  static const standard_13 =
      StandardNumericNumFormat._(numFmtId: 13, formatCode: "# ??/??");
  static const standard_14 =
      StandardDateTimeNumFormat._(numFmtId: 14, formatCode: "mm-dd-yy");
  static const standard_15 =
      StandardDateTimeNumFormat._(numFmtId: 15, formatCode: "d-mmm-yy");
  static const standard_16 =
      StandardDateTimeNumFormat._(numFmtId: 16, formatCode: "d-mmm");
  static const standard_17 =
      StandardDateTimeNumFormat._(numFmtId: 17, formatCode: "mmm-yy");
  static const standard_18 =
      StandardTimeNumFormat._(numFmtId: 18, formatCode: "h:mm AM/PM");
  static const standard_19 =
      StandardTimeNumFormat._(numFmtId: 19, formatCode: "h:mm:ss AM/PM");
  static const standard_20 =
      StandardTimeNumFormat._(numFmtId: 20, formatCode: "h:mm");
  static const standard_21 =
      StandardTimeNumFormat._(numFmtId: 21, formatCode: "h:mm:dd");
  static const standard_22 =
      StandardDateTimeNumFormat._(numFmtId: 22, formatCode: "m/d/yy h:mm");
  static const standard_37 =
      StandardNumericNumFormat._(numFmtId: 37, formatCode: "#,##0 ;(#,##0)");
  static const standard_38 = StandardNumericNumFormat._(
      numFmtId: 38, formatCode: "#,##0 ;[Red](#,##0)");
  static const standard_39 = StandardNumericNumFormat._(
      numFmtId: 39, formatCode: "#,##0.00;(#,##0.00)");
  static const standard_40 = StandardNumericNumFormat._(
      numFmtId: 40, formatCode: "#,##0.00;[Red](#,#)");
  static const standard_45 =
      StandardTimeNumFormat._(numFmtId: 45, formatCode: "mm:ss");
  static const standard_46 =
      StandardTimeNumFormat._(numFmtId: 46, formatCode: "[h]:mm:ss");
  static const standard_47 =
      StandardTimeNumFormat._(numFmtId: 47, formatCode: "mmss.0");
  static const standard_48 =
      StandardNumericNumFormat._(numFmtId: 48, formatCode: "##0.0");
  static const standard_49 =
      StandardNumericNumFormat._(numFmtId: 49, formatCode: "@");
}
