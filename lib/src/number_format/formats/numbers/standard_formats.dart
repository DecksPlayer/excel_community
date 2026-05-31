part of '../../../../excel_community.dart';

const Map<int, NumFormat> _standardNumFormats = {
  0: NumFormat.standard_0,
  1: NumFormat.standard_1,
  2: NumFormat.standard_2,
  3: NumFormat.standard_3,
  4: NumFormat.standard_4,
  // Currency (5–8)
  5: NumFormat.standard_5,
  6: NumFormat.standard_6,
  7: NumFormat.standard_7,
  8: NumFormat.standard_8,
  9: NumFormat.standard_9,
  10: NumFormat.standard_10,
  11: NumFormat.standard_11,
  12: NumFormat.standard_12,
  13: NumFormat.standard_13,
  // Date / Time (14–22)
  14: NumFormat.standard_14,
  15: NumFormat.standard_15,
  16: NumFormat.standard_16,
  17: NumFormat.standard_17,
  18: NumFormat.standard_18,
  19: NumFormat.standard_19,
  20: NumFormat.standard_20,
  21: NumFormat.standard_21,
  22: NumFormat.standard_22,
  // Reserved (23–26) — not in OOXML spec, treated as General
  23: NumFormat.standard_23,
  24: NumFormat.standard_24,
  25: NumFormat.standard_25,
  26: NumFormat.standard_26,
  // CJK locale Date/Time (27–36)
  27: NumFormat.standard_27,
  28: NumFormat.standard_28,
  29: NumFormat.standard_29,
  30: NumFormat.standard_30,
  31: NumFormat.standard_31,
  32: NumFormat.standard_32,
  33: NumFormat.standard_33,
  34: NumFormat.standard_34,
  35: NumFormat.standard_35,
  36: NumFormat.standard_36,
  // Accounting without fill (37–40)
  37: NumFormat.standard_37,
  38: NumFormat.standard_38,
  39: NumFormat.standard_39,
  40: NumFormat.standard_40,
  // Accounting with fill character (41–44)
  41: NumFormat.standard_41,
  42: NumFormat.standard_42,
  43: NumFormat.standard_43,
  44: NumFormat.standard_44,
  // Time (45–47)
  45: NumFormat.standard_45,
  46: NumFormat.standard_46,
  47: NumFormat.standard_47,
  // Numeric / Text (48–49)
  48: NumFormat.standard_48,
  49: NumFormat.standard_49,
};

/// Helper to keep NumFormat clean.
/// Format codes follow ECMA-376 §18.8.30 and Microsoft's built-in format table.
class StandardFormats {
  // --- General / Numeric (0–4) ---
  static const standard_0 =
      StandardNumericNumFormat._(numFmtId: 0, formatCode: 'General');
  static const standard_1 =
      StandardNumericNumFormat._(numFmtId: 1, formatCode: '0');
  static const standard_2 =
      StandardNumericNumFormat._(numFmtId: 2, formatCode: '0.00');
  static const standard_3 =
      StandardNumericNumFormat._(numFmtId: 3, formatCode: '#,##0');
  static const standard_4 =
      StandardNumericNumFormat._(numFmtId: 4, formatCode: '#,##0.00');

  // --- Currency (5–8) ---
  static const standard_5 = StandardNumericNumFormat._(
      numFmtId: 5, formatCode: r'$#,##0_);($#,##0)');
  static const standard_6 = StandardNumericNumFormat._(
      numFmtId: 6, formatCode: r'$#,##0_);[Red]($#,##0)');
  static const standard_7 = StandardNumericNumFormat._(
      numFmtId: 7, formatCode: r'$#,##0.00_);($#,##0.00)');
  static const standard_8 = StandardNumericNumFormat._(
      numFmtId: 8, formatCode: r'$#,##0.00_);[Red]($#,##0.00)');

  // --- Percentage / Scientific / Fraction (9–13) ---
  static const standard_9 =
      StandardNumericNumFormat._(numFmtId: 9, formatCode: '0%');
  static const standard_10 =
      StandardNumericNumFormat._(numFmtId: 10, formatCode: '0.00%');
  static const standard_11 =
      StandardNumericNumFormat._(numFmtId: 11, formatCode: '0.00E+00');
  static const standard_12 =
      StandardNumericNumFormat._(numFmtId: 12, formatCode: '# ?/?');
  static const standard_13 =
      StandardNumericNumFormat._(numFmtId: 13, formatCode: '# ??/??');

  // --- Date / Time (14–22) ---
  static const standard_14 =
      StandardDateTimeNumFormat._(numFmtId: 14, formatCode: 'mm-dd-yy');
  static const standard_15 =
      StandardDateTimeNumFormat._(numFmtId: 15, formatCode: 'd-mmm-yy');
  static const standard_16 =
      StandardDateTimeNumFormat._(numFmtId: 16, formatCode: 'd-mmm');
  static const standard_17 =
      StandardDateTimeNumFormat._(numFmtId: 17, formatCode: 'mmm-yy');
  static const standard_18 =
      StandardTimeNumFormat._(numFmtId: 18, formatCode: 'h:mm AM/PM');
  static const standard_19 =
      StandardTimeNumFormat._(numFmtId: 19, formatCode: 'h:mm:ss AM/PM');
  static const standard_20 =
      StandardTimeNumFormat._(numFmtId: 20, formatCode: 'h:mm');
  static const standard_21 =
      StandardTimeNumFormat._(numFmtId: 21, formatCode: 'h:mm:ss'); // fixed: was 'h:mm:dd'
  static const standard_22 =
      StandardDateTimeNumFormat._(numFmtId: 22, formatCode: 'm/d/yy h:mm');

  // --- Reserved (23–26) — not defined in OOXML spec, fallback to General ---
  static const standard_23 =
      StandardNumericNumFormat._(numFmtId: 23, formatCode: 'General');
  static const standard_24 =
      StandardNumericNumFormat._(numFmtId: 24, formatCode: 'General');
  static const standard_25 =
      StandardNumericNumFormat._(numFmtId: 25, formatCode: 'General');
  static const standard_26 =
      StandardNumericNumFormat._(numFmtId: 26, formatCode: 'General');

  // --- CJK locale Date/Time (27–36) ---
  // Note: some share identical formatCodes but must be separate const instances
  // so the inverse map (NumFormat→ID) remains injective.
  static const standard_27 = StandardDateTimeNumFormat._(  // Chinese Traditional e/m/d
      numFmtId: 27, formatCode: r'[$-404]e/m/d');
  static const standard_28 = StandardDateTimeNumFormat._(  // Chinese Traditional am/pm
      numFmtId: 28, formatCode: r'[$-404]e/m/d h:mm AM/PM');
  static const standard_29 = StandardDateTimeNumFormat._(  // Chinese Traditional e/m/d (2)
      numFmtId: 29, formatCode: r'[$-404]e"年"m"月"d"日"');
  static const standard_30 =
      StandardDateTimeNumFormat._(numFmtId: 30, formatCode: 'm/d/yy');
  static const standard_31 = StandardDateTimeNumFormat._(  // Chinese Simplified
      numFmtId: 31, formatCode: 'yyyy"年"m"月"d"日"');
  static const standard_32 =
      StandardTimeNumFormat._(numFmtId: 32, formatCode: 'h"時"mm"分"');
  static const standard_33 =
      StandardTimeNumFormat._(numFmtId: 33, formatCode: 'h"時"mm"分"ss"秒"');
  static const standard_34 =
      StandardTimeNumFormat._(numFmtId: 34, formatCode: '上午/下午h"時"mm"分"');
  static const standard_35 =
      StandardTimeNumFormat._(numFmtId: 35, formatCode: '上午/下午h"時"mm"分"ss"秒"');
  static const standard_36 = StandardDateTimeNumFormat._(  // Chinese Traditional e/m/d (3)
      numFmtId: 36, formatCode: r'[$-404]e"月"m"日"d"日"');

  // --- Accounting without fill character (37–40) ---
  static const standard_37 =
      StandardNumericNumFormat._(numFmtId: 37, formatCode: '#,##0 ;(#,##0)');
  static const standard_38 = StandardNumericNumFormat._(
      numFmtId: 38, formatCode: '#,##0 ;[Red](#,##0)');
  static const standard_39 = StandardNumericNumFormat._(
      numFmtId: 39, formatCode: '#,##0.00;(#,##0.00)');
  static const standard_40 = StandardNumericNumFormat._(
      numFmtId: 40,
      formatCode: '#,##0.00;[Red](#,##0.00)'); // fixed: was '#,##0.00;[Red](#,#)'

  // --- Accounting with fill character (41–44) ---
  static const standard_41 = StandardNumericNumFormat._(
      numFmtId: 41,
      formatCode: r'_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)');
  static const standard_42 = StandardNumericNumFormat._(
      numFmtId: 42,
      formatCode: r'_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)');
  static const standard_43 = StandardNumericNumFormat._(
      numFmtId: 43,
      formatCode: r'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)');
  static const standard_44 = StandardNumericNumFormat._(
      numFmtId: 44,
      formatCode: r'_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');

  // --- Time (45–47) ---
  static const standard_45 =
      StandardTimeNumFormat._(numFmtId: 45, formatCode: 'mm:ss');
  static const standard_46 =
      StandardTimeNumFormat._(numFmtId: 46, formatCode: '[h]:mm:ss');
  static const standard_47 =
      StandardTimeNumFormat._(numFmtId: 47, formatCode: 'mmss.0');

  // --- Numeric / Text (48–49) ---
  static const standard_48 =
      StandardNumericNumFormat._(numFmtId: 48, formatCode: '##0.0E+0'); // fixed: was '##0.0'
  static const standard_49 =
      StandardNumericNumFormat._(numFmtId: 49, formatCode: '@');
}
