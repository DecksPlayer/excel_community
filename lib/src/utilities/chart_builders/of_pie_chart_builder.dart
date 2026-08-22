part of '../../../excel_community.dart';

/// Builder for OfPie (Pie-of-Pie and Bar-of-Pie) chart styles.
class OfPieChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    final ofPie = chart as OfPieChart;
    builder.element('c:ofPieType',
        attributes: {'val': ofPie.ofPieType.ooxmlValue});
    builder.element('c:splitType',
        attributes: {'val': ofPie.splitType.ooxmlValue});
    if (ofPie.splitPosition != null) {
      builder.element('c:splitPos',
          attributes: {'val': ofPie.splitPosition.toString()});
    }
    builder.element('c:secondPieSize',
        attributes: {'val': ofPie.secondPieSize.toString()});
    builder.element('c:serLines');
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

            final borderColor =
                userStyle.borderColor ?? userStyle.fillColor!;
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
      final colors = ChartColorConfig.getRandomizedPieColors(valuesCount);
      for (int i = 0; i < valuesCount; i++) {
        builder.element('c:dPt', nest: () {
          builder.element('c:idx', attributes: {'val': '$i'});
          builder.element('c:spPr', nest: () {
            builder.element('a:solidFill', nest: () {
              builder.element(
                  'a:srgbClr', attributes: {'val': colors[i].colorHex6});
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
