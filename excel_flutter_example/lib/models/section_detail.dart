import 'package:flutter/material.dart';

enum SelectedSection {
  simpleExcel,
  columnChart,
  lineChart,
  pieChart,
  areaChart,
  doughnutChart,
  radarChart,
  barChart,
  scatterChart,
  imageEmbedding,
  textStyles,
  numberFormats,
  multiSheets,
  allCharts,
  fullDemo,
}

enum ChartType { column, line, pie, area, doughnut, radar, bar, scatter }

class SectionDetail {
  final String title;
  final String description;
  final IconData icon;
  final Color themeColor;
  final List<String> highlights;
  final String codeSnippet;

  SectionDetail({
    required this.title,
    required this.description,
    required this.icon,
    required this.themeColor,
    required this.highlights,
    required this.codeSnippet,
  });
}
