part of '../../excel_community.dart';

/// Configuration for chart colors and palettes.
///
/// Provides centralized color management for all chart types with
/// distinct palettes optimized for different chart categories.
class ChartColorConfig {
  ChartColorConfig._();

  // ========================================================================
  // COLOR PALETTES
  // ========================================================================

  /// Color palette for series-based charts (Column, Bar, Line, Area, Scatter)
  /// 12 vibrant, distinguishable colors
  static final List<ExcelColor> seriesPalette = [
    ExcelColor.fromHexString('4472C4'), // Blue
    ExcelColor.fromHexString('ED7D31'), // Orange
    ExcelColor.fromHexString('70AD47'), // Green
    ExcelColor.fromHexString('FFC000'), // Gold
    ExcelColor.fromHexString('5B9BD5'), // Light Blue
    ExcelColor.fromHexString('C5504B'), // Red
    ExcelColor.fromHexString('8064A2'), // Purple
    ExcelColor.fromHexString('4BACC6'), // Cyan
    ExcelColor.fromHexString('9BBB59'), // Olive
    ExcelColor.fromHexString('F79646'), // Light Orange
    ExcelColor.fromHexString('17B897'), // Teal
    ExcelColor.fromHexString('E83352'), // Crimson
  ];

  /// Color palette for radar charts
  /// 8 colors optimized for multi-dimensional comparison
  static final List<ExcelColor> radarPalette = [
    ExcelColor.fromHexString('4472C4'), // Blue
    ExcelColor.fromHexString('ED7D31'), // Orange
    ExcelColor.fromHexString('70AD47'), // Green
    ExcelColor.fromHexString('FFC000'), // Gold
    ExcelColor.fromHexString('5B9BD5'), // Light Blue
    ExcelColor.fromHexString('C5504B'), // Red
    ExcelColor.fromHexString('8064A2'), // Purple
    ExcelColor.fromHexString('4BACC6'), // Cyan
  ];

  /// Color palette for pie/doughnut charts
  /// 20 colors for maximum segment variety with random distribution
  static final List<ExcelColor> piePalette = [
    ExcelColor.fromHexString('4472C4'), // Blue
    ExcelColor.fromHexString('ED7D31'), // Orange
    ExcelColor.fromHexString('A5A5A5'), // Gray
    ExcelColor.fromHexString('FFC000'), // Gold
    ExcelColor.fromHexString('5B9BD5'), // Light Blue
    ExcelColor.fromHexString('70AD47'), // Green
    ExcelColor.fromHexString('264478'), // Dark Blue
    ExcelColor.fromHexString('9E480E'), // Brown
    ExcelColor.fromHexString('636363'), // Dark Gray
    ExcelColor.fromHexString('997300'), // Dark Gold
    ExcelColor.fromHexString('255E91'), // Navy Blue
    ExcelColor.fromHexString('43682B'), // Dark Green
    ExcelColor.fromHexString('C5504B'), // Red
    ExcelColor.fromHexString('8064A2'), // Purple
    ExcelColor.fromHexString('4BACC6'), // Cyan
    ExcelColor.fromHexString('F79646'), // Light Orange
    ExcelColor.fromHexString('9BBB59'), // Olive
    ExcelColor.fromHexString('E83352'), // Crimson
    ExcelColor.fromHexString('17B897'), // Teal
    ExcelColor.fromHexString('FF6F61'), // Coral
  ];

  // ========================================================================
  // COLOR ACCESSORS
  // ========================================================================

  /// Gets a color from the series palette by index (with rotation)
  static ExcelColor getSeriesColor(int index) {
    return seriesPalette[index % seriesPalette.length];
  }

  /// Gets a color from the radar palette by index (with rotation)
  static ExcelColor getRadarColor(int index) {
    return radarPalette[index % radarPalette.length];
  }

  /// Gets a randomized list of pie colors for the specified number of points
  static List<ExcelColor> getRandomizedPieColors(int numPoints) {
    final shuffled = List<ExcelColor>.from(piePalette)..shuffle();
    return shuffled.take(numPoints).toList();
  }

  // ========================================================================
  // STYLE CONSTANTS
  // ========================================================================

  /// Line width for thin borders (Column/Bar chart borders)
  static const String thinLineWidth = '9525'; // ~0.75 pt

  /// Line width for thick lines (Line/Area/Radar charts)
  static const String thickLineWidth = '28575'; // ~2.25 pt

  /// Opacity values (0-100000 scale)
  static const String opacity100 = '100000'; // Fully opaque
  static const String opacity90 = '90000';
  static const String opacity85 = '85000';
  static const String opacity50 = '50000';
  static const String opacity45 = '45000';

  /// White color for borders/backgrounds
  static const ExcelColor white = ExcelColor.white;

  /// Marker sizes
  static const String smallMarker = '5';
  static const String largeMarker = '7';

  // ========================================================================
  // STYLE RESOLUTION HELPERS
  // ========================================================================

  /// Returns the fill color for a series, respecting any user-provided style.
  static ExcelColor resolveFillColor(
      ChartSeries series, int index, List<ExcelColor> palette) {
    return series.style?.fillColor ?? palette[index % palette.length];
  }

  /// Returns the border color for a series (defaults to the fill color when
  /// the user hasn't set a border color explicitly).
  static ExcelColor resolveBorderColor(
      ChartSeries series, int index, List<ExcelColor> palette) {
    final fill = resolveFillColor(series, index, palette);
    return series.style?.borderColor ?? fill;
  }

  // ========================================================================
  // XML HELPER — shared solid-fill emitter
  // ========================================================================

  /// Emits an `<a:solidFill>` element.
  ///
  /// When [alpha] is `null` or 100, no `<a:alpha>` child is added (= fully
  /// opaque).  Any other value emits the OOXML-scale alpha (0–100 000).
  static void emitSolidFill(
      XmlBuilder b, ExcelColor color, [int alpha = 100]) {
    b.element('a:solidFill', nest: () {
      if (alpha >= 100) {
        b.element('a:srgbClr', attributes: {'val': color.colorHex6});
      } else {
        b.element('a:srgbClr', attributes: {'val': color.colorHex6}, nest: () {
          b.element('a:alpha',
              attributes: {'val': ChartSeriesStyle.alphaToOoxml(alpha)});
        });
      }
    });
  }

  /// Builds a complete `<c:spPr>` element for a series, taking the user's
  /// [ChartSeriesStyle] into account when present.
  ///
  /// Parameters:
  /// - [defaultFillAlpha]: opacity % used when there is no user style and the
  ///   chart type renders a transparent fill (area, radar-filled).
  ///   Pass `100` for solid fills (column, bar, scatter).
  /// - [defaultBorderAlpha]: opacity % for the auto-palette border.
  /// - [defaultBorderWidth]: OOXML EMU string for the border line width.
  /// - [includeFill]: set to `false` for line-only charts (plain line chart
  ///   body — fill handled elsewhere).
  static void buildSpPr(
    XmlBuilder b,
    ChartSeries series,
    int seriesIndex,
    List<ExcelColor> palette, {
    int defaultFillAlpha = 100,
    int defaultBorderAlpha = 100,
    String defaultBorderWidth = thinLineWidth,
    bool includeFill = true,
  }) {
    final userStyle = series.style;
    final fillColor = resolveFillColor(series, seriesIndex, palette);
    final borderColor = resolveBorderColor(series, seriesIndex, palette);

    // Resolve effective alphas
    int effectiveFillAlpha;
    bool noFill;
    if (userStyle != null) {
      noFill = userStyle.fillType == ChartFillType.none;
      effectiveFillAlpha = userStyle.fillType == ChartFillType.transparent
          ? userStyle.fillAlpha
          : 100;
    } else {
      noFill = false;
      effectiveFillAlpha = defaultFillAlpha;
    }

    final effectiveBorderAlpha =
        userStyle?.borderAlpha ?? defaultBorderAlpha;
    final effectiveBorderWidth =
        userStyle?.borderWidth ?? defaultBorderWidth;

    b.element('c:spPr', nest: () {
      if (includeFill) {
        if (noFill) {
          b.element('a:noFill');
        } else {
          emitSolidFill(b, fillColor, effectiveFillAlpha);
        }
      }
      b.element('a:ln', attributes: {'w': effectiveBorderWidth}, nest: () {
        emitSolidFill(b, borderColor, effectiveBorderAlpha);
      });
    });
  }
}
