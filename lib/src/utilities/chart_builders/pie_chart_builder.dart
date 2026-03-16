part of excel_community;

/// Builder for Pie and Doughnut chart styles
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
  void buildSeriesStyle(XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    // Parse the range to determine number of data points
    final rangeMatch = RegExp(r'\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)').firstMatch(series.valuesRange);
    if (rangeMatch == null) return;
    
    final startRow = int.parse(rangeMatch.group(2)!);
    final endRow = int.parse(rangeMatch.group(4)!);
    final valuesCount = endRow - startRow + 1;
    
    // Get randomized colors
    final colors = ChartColorConfig.getRandomizedPieColors(valuesCount);
    
    for (int i = 0; i < valuesCount; i++) {
      builder.element('c:dPt', nest: () {
        builder.element('c:idx', attributes: {'val': '$i'});
        builder.element('c:spPr', nest: () {
          builder.element('a:solidFill', nest: () {
            builder.element('a:srgbClr', attributes: {'val': colors[i].colorHex6});
          });
          builder.element('a:ln', attributes: {'w': ChartColorConfig.thinLineWidth}, nest: () {
            builder.element('a:solidFill', nest: () {
              builder.element('a:srgbClr', attributes: {'val': ChartColorConfig.white.colorHex6});
            });
          });
        });
      });
    }
  }
}
