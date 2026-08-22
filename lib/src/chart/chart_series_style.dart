part of '../../excel_community.dart';

/// Fill style for a chart series.
enum ChartFillType {
  /// Completely opaque solid fill (alpha = 100%).
  solid,

  /// Solid fill with transparency — same color but partially see-through.
  /// The transparency level is controlled by [ChartSeriesStyle.fillAlpha].
  transparent,

  /// No fill at all — the series area is left empty (line/border only).
  none,
}

/// Per-series visual styling for charts.
///
/// Attach a [ChartSeriesStyle] to [ChartSeries.style] to override the
/// automatic palette color. All fields are optional: supply only what you
/// want to change and the builder will fill in sensible defaults for the rest.
///
/// ### Fill behavior
/// | [fillType]         | What is rendered |
/// |--------------------|-----------------|
/// | `solid`            | 100 % opaque fill |
/// | `transparent`      | fill with [fillAlpha] % opacity (0 = invisible, 100 = solid) |
/// | `none`             | no fill; only the border line is drawn |
///
/// ### Example — solid custom color
/// ```dart
/// ChartSeries(
///   name: 'Revenue',
///   categoriesRange: r"Sheet1!$A$2:$A$7",
///   valuesRange: r"Sheet1!$B$2:$B$7",
///   style: ChartSeriesStyle(
///     fillColor: ExcelColor.fromHexString('4472C4'),
///     fillType: ChartFillType.solid,
///   ),
/// )
/// ```
///
/// ### Example — semi-transparent + custom border
/// ```dart
/// ChartSeries(
///   name: 'Cost',
///   categoriesRange: r"Sheet1!$A$2:$A$7",
///   valuesRange: r"Sheet1!$C$2:$C$7",
///   style: ChartSeriesStyle(
///     fillColor: ExcelColor.fromHexString('ED7D31'),
///     fillType: ChartFillType.transparent,
///     fillAlpha: 50,                         // 50 % opacity
///     borderColor: ExcelColor.fromHexString('C55A11'),
///   ),
/// )
/// ```
class ChartSeriesStyle {
  /// Fill color of the series.
  ///
  /// When `null` the builder uses the automatic palette color.
  final ExcelColor? fillColor;

  /// Fill style. Defaults to [ChartFillType.solid].
  final ChartFillType fillType;

  /// Opacity of the fill, expressed as a percentage from 0 (fully transparent)
  /// to 100 (fully opaque). Only used when [fillType] is
  /// [ChartFillType.transparent]. Defaults to 50.
  final int fillAlpha;

  /// Border / line color. When `null` the fill color is reused for the border.
  final ExcelColor? borderColor;

  /// Opacity of the border, 0–100. Defaults to 100 (fully opaque).
  final int borderAlpha;

  /// Width of the border line in EMUs (English Metric Units).
  ///
  /// Common values:
  /// - `9525`  → ~0.75 pt (thin, used for Column/Bar borders)
  /// - `19050` → ~1.5 pt  (medium)
  /// - `28575` → ~2.25 pt (thick, used for Line/Area/Radar)
  ///
  /// When `null`, each chart type uses its own default width.
  final String? borderWidth;

  const ChartSeriesStyle({
    this.fillColor,
    this.fillType = ChartFillType.solid,
    this.fillAlpha = 50,
    this.borderColor,
    this.borderAlpha = 100,
    this.borderWidth,
  })  : assert(fillAlpha >= 0 && fillAlpha <= 100,
            'fillAlpha must be between 0 and 100'),
        assert(borderAlpha >= 0 && borderAlpha <= 100,
            'borderAlpha must be between 0 and 100');

  /// Converts a 0–100 percentage into the OOXML alpha scale (0–100 000).
  static String alphaToOoxml(int percent) =>
      (percent * 1000).clamp(0, 100000).toString();
}
