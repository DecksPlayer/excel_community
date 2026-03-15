part of '../../../excel_community.dart';

class _StyleManager {
  final Excel _excel;
  final Save _save;
  late final _StyleResourceCollector _collector;
  late final _StyleXmlBuilders _builders;

  _StyleManager(this._excel, this._save) {
    _collector = _StyleResourceCollector(_excel, _save);
    _builders = _StyleXmlBuilders(_excel, _save);
  }

  void processStylesFile() {
    // 1. Collect resources
    final resources = _collector.collect();
    
    // Update _save._innerCellStyle for other parts of the system that might rely on it
    _save._innerCellStyle.clear();
    _save._innerCellStyle.addAll(resources.innerCellStyle);

    final styleDoc = _excel._xmlFiles['xl/styles.xml']!;

    // 2. Build individual sections using builders
    _builders.buildFonts(
      styleDoc.findAllElements('fonts').first,
      resources.innerFontStyle,
    );

    _builders.buildFills(
      styleDoc.findAllElements('fills').first,
      resources.innerPatternFill,
    );

    _builders.buildBorders(
      styleDoc.findAllElements('borders').first,
      resources.innerBorderSet,
    );

    _builders.buildCellXfs(
      styleDoc.findAllElements('cellXfs').first,
      resources.innerCellStyle,
      resources.innerFontStyle,
      resources.innerPatternFill,
      resources.innerBorderSet,
    );

    _builders.buildNumFmts(styleDoc);
  }
}
