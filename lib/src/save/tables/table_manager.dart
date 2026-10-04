part of '../../../excel_community.dart';

const _tableRelationshipType =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
const _tableContentType =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

/// Writes Excel tables as `xl/tables/tableN.xml` parts.
///
/// Table parts are rebuilt from the model on every save: the ones read from
/// the original file (and their relationships and content types) are
/// dropped, so removed tables leave nothing behind.
class _TableManager {
  final Excel _excel;
  final Save _save;

  _TableManager(this._excel, this._save);

  /// Makes table names unique in the workbook and refreshes header and
  /// totals cells. Runs before styles are collected, since it writes cells.
  void syncTables() {
    final used = <String>{};
    for (final sheet in _excel._sheetMap.values) {
      for (var i = 0; i < sheet._tables.length; i++) {
        var table = sheet._tables[i];
        var name = table.name;
        var n = 2;
        while (!used.add(name.toLowerCase())) {
          name = '${table.name}${n++}';
        }
        if (name != table.name) {
          table = table.copyWith(name: name);
          sheet._tables[i] = table;
        }
        sheet._syncTableCells(table);
      }
    }
  }

  void processTables() {
    // Drop table parts from the original file and from earlier saves.
    _excel._xmlFiles['[Content_Types].xml']?.rootElement.children.removeWhere(
        (node) => node is XmlElement && node.getAttribute('ContentType') == _tableContentType);
    for (final file in _excel._archive.files) {
      if (file.name.startsWith('xl/tables/')) _save._excludedParts.add(file.name);
    }
    _excel._xmlFiles.removeWhere((path, _) => path.startsWith('xl/tables/'));

    var id = 0;
    _excel._sheetMap.forEach((sheetName, sheet) {
      sheet._tableRIds.clear();
      final sheetId = _excel._xmlSheetId[sheetName];
      if (sheetId == null) return;
      final relsPath = 'xl/worksheets/_rels/${sheetId.split('/').last}.rels';
      _save._loadXmlPart(relsPath)?.rootElement.children.removeWhere((node) =>
          node is XmlElement && node.getAttribute('Type') == _tableRelationshipType);
      if (sheet._tables.isEmpty) return;

      final relationships = _save._relationshipsPart(relsPath).rootElement;
      for (final table in sheet._tables) {
        id++;
        final path = 'xl/tables/table$id.xml';
        _excel._xmlFiles[path] = XmlDocument.parse(table._toXmlString(id));
        _save._addContentType(_tableContentType, '/$path');
        final rId = _save._nextRelationshipId(relationships);
        relationships.children.add(XmlElement(XmlName('Relationship'), [
          XmlAttribute(XmlName('Id'), rId),
          XmlAttribute(XmlName('Type'), _tableRelationshipType),
          XmlAttribute(XmlName('Target'), '../tables/table$id.xml'),
        ]));
        sheet._tableRIds.add(rId);
      }
    });
  }
}
