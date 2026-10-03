part of '../../../excel_community.dart';

const _hyperlinkRelationshipType =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

/// Registers external hyperlinks as relationships of each worksheet
/// (`TargetMode="External"`) before the worksheet XML is written.
class _HyperlinkManager {
  final Excel _excel;
  final Save _save;

  _HyperlinkManager(this._excel, this._save);

  void processHyperlinks() {
    _excel._sheetMap.forEach((sheetName, sheet) {
      sheet._hyperlinkRIds.clear();
      final sheetId = _excel._xmlSheetId[sheetName];
      if (sheetId == null) return;
      final sheetFileName = sheetId.split('/').last;
      final relsPath = 'xl/worksheets/_rels/$sheetFileName.rels';

      // Hyperlink relationships are rebuilt from the model, so drop the
      // ones read from the file (other relationships are kept).
      final existing = _save._loadXmlPart(relsPath);
      existing?.rootElement.children.removeWhere((node) =>
          node is XmlElement &&
          node.getAttribute('Type') == _hyperlinkRelationshipType);

      final external =
          sheet._hyperlinks.entries.where((e) => e.value.isExternal);
      if (external.isEmpty) return;

      final relationships = _save._relationshipsPart(relsPath).rootElement;
      for (final entry in external) {
        final rId = _save._nextRelationshipId(relationships);
        relationships.children.add(XmlElement(XmlName('Relationship'), [
          XmlAttribute(XmlName('Id'), rId),
          XmlAttribute(XmlName('Type'), _hyperlinkRelationshipType),
          XmlAttribute(XmlName('Target'), entry.value.url!),
          XmlAttribute(XmlName('TargetMode'), 'External'),
        ]));
        sheet._hyperlinkRIds[entry.key] = rId;
      }
    });
  }
}
