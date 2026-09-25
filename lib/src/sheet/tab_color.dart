part of '../../excel_community.dart';

/// Normalizes any color hex string (#RRGGBB, RRGGBB, #AARRGGBB, AARRGGBB)
/// into an 8-character uppercase ARGB string (AARRGGBB) suitable for OpenXML.
String _normalizeColorHex(String hex) {
  var clean = hex.replaceAll('#', '').trim().toUpperCase();
  if (clean.length == 6) {
    return 'FF$clean';
  } else if (clean.length == 8) {
    return clean;
  } else if (clean.length == 3) {
    final r = clean[0];
    final g = clean[1];
    final b = clean[2];
    return 'FF$r$r$g$g$b$b';
  }
  return clean;
}

/// Represents the OpenXML `<tabColor>` element for a worksheet's tab.
///
/// Used inside `<sheetPr>` to assign custom colors to the sheet tab shown at the
/// bottom of spreadsheet applications like Microsoft Excel, LibreOffice Calc,
/// and Google Sheets.
class TabColor extends Equatable {
  /// The [ExcelColor] associated with this tab color, if created from or resolved to one.
  final ExcelColor? color;

  /// 8-character ARGB hex string (e.g. `'FFFF0000'` for red).
  final String? rgb;

  /// Optional 0-based theme color index.
  final int? theme;

  /// Optional tint value applied to the theme color (-1.0 to 1.0).
  final double? tint;

  /// Optional legacy indexed color value.
  final int? indexed;

  /// Optional auto color flag.
  final bool? auto;

  const TabColor({
    this.color,
    this.rgb,
    this.theme,
    this.tint,
    this.indexed,
    this.auto,
  });

  /// Creates a [TabColor] from an [ExcelColor].
  factory TabColor.fromColor(ExcelColor color) {
    final normalized = color == ExcelColor.none ? null : _normalizeColorHex(color.colorHex);
    return TabColor(
      color: color,
      rgb: normalized,
    );
  }

  /// Creates a [TabColor] from a hex string (e.g. `'#FF0000'`, `'4CAF50'`, or `'FF4CAF50'`).
  factory TabColor.fromHex(String hex) {
    final normalized = _normalizeColorHex(hex);
    return TabColor(
      color: ExcelColor.fromHexString(normalized),
      rgb: normalized,
    );
  }

  /// Creates a [TabColor] from a theme color index with an optional [tint].
  factory TabColor.fromTheme(int theme, {double? tint}) {
    return TabColor(theme: theme, tint: tint);
  }

  /// Returns the effective 8-character ARGB hex string, or `null` if theme/auto based.
  String? get colorHex =>
      rgb ?? (color != null && color != ExcelColor.none ? _normalizeColorHex(color!.colorHex) : null);

  /// Returns a 6-character hex string (`RRGGBB`), or `null` if not available.
  String? get colorHex6 {
    final hex = colorHex;
    if (hex == null || hex == 'NONE') return null;
    if (hex.length >= 6) {
      return hex.substring(hex.length - 6);
    }
    return hex.padLeft(6, '0');
  }

  TabColor copyWith({
    ExcelColor? color,
    String? rgb,
    int? theme,
    double? tint,
    int? indexed,
    bool? auto,
  }) {
    return TabColor(
      color: color ?? this.color,
      rgb: rgb ?? this.rgb,
      theme: theme ?? this.theme,
      tint: tint ?? this.tint,
      indexed: indexed ?? this.indexed,
      auto: auto ?? this.auto,
    );
  }

  /// Serializes this [TabColor] to an OpenXML `<tabColor .../>` element string.
  String toXmlString() {
    final sb = StringBuffer('<tabColor');
    final effectiveRgb = colorHex;
    if (effectiveRgb != null && effectiveRgb.isNotEmpty && effectiveRgb != 'NONE') {
      sb.write(' rgb="$effectiveRgb"');
    }
    if (theme != null) {
      sb.write(' theme="$theme"');
    }
    if (tint != null) {
      sb.write(' tint="$tint"');
    }
    if (indexed != null) {
      sb.write(' indexed="$indexed"');
    }
    if (auto != null) {
      sb.write(' auto="${auto! ? '1' : '0'}"');
    }
    sb.write('/>');
    return sb.toString();
  }

  @override
  List<Object?> get props => [
        colorHex,
        theme,
        tint,
        indexed,
        auto,
      ];
}
