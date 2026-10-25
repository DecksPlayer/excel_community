part of '../../../excel_community.dart';

/// Builder for Pie and Doughnut chart styles.
///
/// When the user provides a [ChartSeriesStyle] on the single pie series, that
/// color (and fill type) is applied to **all** slices uniformly. Without a
/// custom style the builder uses the automatic shuffled 20-color palette.
class PieChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    if (chart is PieChart) {
      builder.element('c:firstSliceAng', attributes: {'val': '0'});
    } else if (chart is DoughnutChart) {
      builder.element('c:holeSize', attributes: {'val': '50'});
    }
  }

  @override
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final rangeMatch = RegExp(r'\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)')
        .firstMatch(series.valuesRange);
    if (rangeMatch == null) return;

    final startRow = int.parse(rangeMatch.group(2)!);
    final endRow = int.parse(rangeMatch.group(4)!);
    final valuesCount = endRow - startRow + 1;

    final userStyle = series.style;

    if (userStyle?.fillColor != null) {
      // User-defined: apply same style to all slices.
      for (int i = 0; i < valuesCount; i++) {
        builder.element('c:dPt', nest: () {
          builder.element('c:idx', attributes: {'val': '$i'});
          builder.element('c:spPr', nest: () {
            final noFill = userStyle!.fillType == ChartFillType.none;
            final alpha = userStyle.fillType == ChartFillType.transparent
                ? userStyle.fillAlpha
                : 100;

            if (noFill) {
              builder.element('a:noFill');
            } else {
              ChartColorConfig.emitSolidFill(
                  builder, userStyle.fillColor!, alpha);
            }

            final borderColor = userStyle.borderColor ?? userStyle.fillColor!;
            final borderWidth =
                userStyle.borderWidth ?? ChartColorConfig.thinLineWidth;
            builder.element('a:ln', attributes: {'w': borderWidth}, nest: () {
              ChartColorConfig.emitSolidFill(
                  builder, borderColor, userStyle.borderAlpha);
            });
          });
        });
      }
    } else {
      // Automatic: palette in order, white separator border.
      for (int i = 0; i < valuesCount; i++) {
        builder.element('c:dPt', nest: () {
          builder.element('c:idx', attributes: {'val': '$i'});
          builder.element('c:spPr', nest: () {
            builder.element('a:solidFill', nest: () {
              builder.element('a:srgbClr', attributes: {
                'val': ChartColorConfig.getPieColor(i).colorHex6
              });
            });
            builder.element('a:ln',
                attributes: {'w': ChartColorConfig.thinLineWidth}, nest: () {
              builder.element('a:solidFill', nest: () {
                builder.element('a:srgbClr',
                    attributes: {'val': ChartColorConfig.white.colorHex6});
              });
            });
          });
        });
      }
    }
  }
}
