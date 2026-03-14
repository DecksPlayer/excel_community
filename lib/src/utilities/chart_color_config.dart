part of excel_community;

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
  static final List<String> seriesPalette = [
    '4472C4', // Blue
    'ED7D31', // Orange
    '70AD47', // Green
    'FFC000', // Gold
    '5B9BD5', // Light Blue
    'C5504B', // Red
    '8064A2', // Purple
    '4BACC6', // Cyan
    '9BBB59', // Olive
    'F79646', // Light Orange
    '17B897', // Teal
    'E83352', // Crimson
  ];

  /// Color palette for radar charts
  /// 8 colors optimized for multi-dimensional comparison
  static final List<String> radarPalette = [
    '4472C4', // Blue
    'ED7D31', // Orange
    '70AD47', // Green
    'FFC000', // Gold
    '5B9BD5', // Light Blue
    'C5504B', // Red
    '8064A2', // Purple
    '4BACC6', // Cyan
  ];

  /// Color palette for pie/doughnut charts
  /// 20 colors for maximum segment variety with random distribution
  static final List<String> piePalette = [
    '4472C4', // Blue
    'ED7D31', // Orange
    'A5A5A5', // Gray
    'FFC000', // Gold
    '5B9BD5', // Light Blue
    '70AD47', // Green
    '264478', // Dark Blue
    '9E480E', // Brown
    '636363', // Dark Gray
    '997300', // Dark Gold
    '255E91', // Navy Blue
    '43682B', // Dark Green
    'C5504B', // Red
    '8064A2', // Purple
    '4BACC6', // Cyan
    'F79646', // Light Orange
    '9BBB59', // Olive
    'E83352', // Crimson
    '17B897', // Teal
    'FF6F61', // Coral
  ];

  // ========================================================================
  // COLOR ACCESSORS
  // ========================================================================

  /// Gets a color from the series palette by index (with rotation)
  static String getSeriesColor(int index) {
    return seriesPalette[index % seriesPalette.length];
  }

  /// Gets a color from the radar palette by index (with rotation)
  static String getRadarColor(int index) {
    return radarPalette[index % radarPalette.length];
  }

  /// Gets a randomized list of pie colors for the specified number of points
  static List<String> getRandomizedPieColors(int numPoints) {
    final shuffled = List<String>.from(piePalette)..shuffle();
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
  static const String white = 'FFFFFF';

  /// Marker sizes
  static const String smallMarker = '5';
  static const String largeMarker = '7';
}
