part of '../../../excel_community.dart';

class _CommentManager {
  final Excel _excel;
  final Save _save;

  _CommentManager(this._excel, this._save);

  void processComments() {
    int commentsCount = 0;
    int vmlCount = 0;

    // Count existing files to avoid overwriting unrelated ones
    for (final filename in _excel._archive.files.map((f) => f.name)) {
      if (filename.startsWith('xl/comments') && filename.endsWith('.xml')) {
        final match = RegExp(r'\d+').stringMatch(filename.split('/').last);
        if (match != null) {
          final idx = int.tryParse(match);
          if (idx != null && idx > commentsCount) {
            commentsCount = idx;
          }
        }
      } else if (filename.startsWith('xl/drawings/vmlDrawing') && filename.endsWith('.vml')) {
        final match = RegExp(r'\d+').stringMatch(filename.split('/').last);
        if (match != null) {
          final idx = int.tryParse(match);
          if (idx != null && idx > vmlCount) {
            vmlCount = idx;
          }
        }
      }
    }

    _excel._sheetMap.forEach((sheetName, sheet) {
      // Find cells with comments
      final cellComments = <String, String>{}; // e.g. {"A1": "My comment"}
      final cellCoords = <(int, int)>[];

      for (final rowIndex in sheet._sheetData.keys) {
        final row = sheet._sheetData[rowIndex]!;
        for (final colIndex in row.keys) {
          final cellData = row[colIndex]!;
          if (cellData.comment != null && cellData.comment!.isNotEmpty) {
            final cellRef = cellData.cellIndex.cellId;
            cellComments[cellRef] = cellData.comment!;
            cellCoords.add((colIndex, rowIndex));
          }
        }
      }

      if (cellComments.isEmpty) {
        return;
      }

      final sheetId = _excel._xmlSheetId[sheetName]!;
      final sheetFileName = sheetId.split('/').last; // e.g. sheet1.xml
      final sheetRelsPath = 'xl/worksheets/_rels/$sheetFileName.rels';

      // Find or create sheet rels document
      var sheetRels = _excel._xmlFiles[sheetRelsPath];
      if (sheetRels == null) {
        final relsBuilder = XmlBuilder();
        relsBuilder.processing(
            'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
        relsBuilder.element('Relationships',
            attributes: {
              'xmlns':
                  'http://schemas.openxmlformats.org/package/2006/relationships',
            },
            nest: () {});
        sheetRels = relsBuilder.buildDocument();
        _excel._xmlFiles[sheetRelsPath] = sheetRels;
      }

      final sheetRelsRoot = sheetRels.findAllElements('Relationships').first;

      var (commentsPath, commentsRId, vmlPath, vmlRId) =
          _findExistingCommentsAndVml(sheetRelsPath);

      if (commentsPath == null) {
        commentsCount++;
        commentsPath = 'xl/comments$commentsCount.xml';
      }
      if (vmlPath == null) {
        vmlCount++;
        vmlPath = 'xl/drawings/vmlDrawing$vmlCount.vml';
      }

      if (commentsRId == null) {
        final nextRIdIndex =
            sheetRelsRoot.children.whereType<XmlElement>().length + 1;
        commentsRId = 'rId$nextRIdIndex';
        
        final commentsFileName = commentsPath.split('/').last;
        sheetRelsRoot.children.add(XmlElement(XmlName.parts('Relationship'), [
          XmlAttribute(XmlName.parts('Id'), commentsRId),
          XmlAttribute(
            XmlName.parts('Type'),
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
          ),
          XmlAttribute(XmlName.parts('Target'), '../$commentsFileName'),
        ]));
      }

      if (vmlRId == null) {
        final nextRIdIndex =
            sheetRelsRoot.children.whereType<XmlElement>().length + 1;
        vmlRId = 'rId$nextRIdIndex';
        
        final vmlFileName = vmlPath.split('/').last;
        sheetRelsRoot.children.add(XmlElement(XmlName.parts('Relationship'), [
          XmlAttribute(XmlName.parts('Id'), vmlRId),
          XmlAttribute(
            XmlName.parts('Type'),
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
          ),
          XmlAttribute(XmlName.parts('Target'), '../drawings/$vmlFileName'),
        ]));
      }

      // Update the legacyDrawing relationship ID in the Sheet
      sheet._legacyDrawingRId = vmlRId;

      // 1. Generate comments XML
      final commentsBuilder = XmlBuilder();
      commentsBuilder.processing(
          'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
      commentsBuilder.element('comments', attributes: {
        'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      }, nest: () {
        commentsBuilder.element('authors', nest: () {
          commentsBuilder.element('author', nest: 'Author');
        });
        commentsBuilder.element('commentList', nest: () {
          for (final entry in cellComments.entries) {
            commentsBuilder.element('comment', attributes: {
              'ref': entry.key,
              'authorId': '0',
            }, nest: () {
              commentsBuilder.element('text', nest: () {
                commentsBuilder.element('r', nest: () {
                  commentsBuilder.element('t', nest: entry.value);
                });
              });
            });
          }
        });
      });

      _excel._xmlFiles[commentsPath] = commentsBuilder.buildDocument();

      // Add to Content Types
      _save._addContentType(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
        '/$commentsPath',
      );

      // 2. Generate VML drawing XML
      final vmlContent = _generateVml(cellCoords);
      _excel._sheetXmls[vmlPath] = vmlContent;

      _save._addDefaultContentType(
        'application/vnd.openxmlformats-officedocument.vmlDrawing',
        'vml',
      );
    });
  }

  (String?, String?, String?, String?) _findExistingCommentsAndVml(String sheetRelsPath) {
    final sheetRels = _excel._xmlFiles[sheetRelsPath];
    if (sheetRels == null) return (null, null, null, null);

    String? commentsPath;
    String? commentsRId;
    String? vmlPath;
    String? vmlRId;

    for (final rel in sheetRels.findAllElements('Relationship')) {
      final type = rel.getAttribute('Type') ?? '';
      final id = rel.getAttribute('Id');
      final target = rel.getAttribute('Target') ?? '';
      if (type == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments') {
        commentsRId = id;
        if (target.startsWith('../')) {
          commentsPath = 'xl/' + target.substring(3);
        } else if (target.startsWith('/')) {
          commentsPath = target.substring(1);
        } else if (!target.startsWith('xl/')) {
          commentsPath = 'xl/worksheets/$target';
        } else {
          commentsPath = target;
        }
      } else if (type == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing') {
        vmlRId = id;
        if (target.startsWith('../')) {
          vmlPath = 'xl/' + target.substring(3);
        } else if (target.startsWith('/')) {
          vmlPath = target.substring(1);
        } else if (!target.startsWith('xl/')) {
          vmlPath = 'xl/worksheets/$target';
        } else {
          vmlPath = target;
        }
      }
    }
    return (commentsPath, commentsRId, vmlPath, vmlRId);
  }

  String _generateVml(List<(int, int)> cellCoordinates) {
    final buffer = StringBuffer();
    buffer.writeln('<xml xmlns:v="urn:schemas-microsoft-com:vml"');
    buffer.writeln(' xmlns:o="urn:schemas-microsoft-com:office:office"');
    buffer.writeln(' xmlns:x="urn:schemas-microsoft-com:office:excel">');

    buffer.writeln(' <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">');
    buffer.writeln('  <v:stroke joinstyle="miter"/>');
    buffer.writeln('  <v:path gradientshapeok="t" o:connecttype="rect"/>');
    buffer.writeln(' </v:shapetype>');

    for (int i = 0; i < cellCoordinates.length; i++) {
      final col = cellCoordinates[i].$1;
      final row = cellCoordinates[i].$2;
      final shapeId = 1025 + i;

      final startCol = col + 1;
      final startRow = row > 0 ? row - 1 : 0;
      final endCol = col + 3;
      final endRow = row + 4;
      final anchor = '$startCol, 15, $startRow, 10, $endCol, 15, $endRow, 10';

      buffer.writeln(' <v:shape id="_x0000_s$shapeId" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">');
      buffer.writeln('  <v:fill color2="#ffffe1"/>');
      buffer.writeln('  <v:shadow on="t" color="black" obscured="t"/>');
      buffer.writeln('  <v:path o:connecttype="none"/>');
      buffer.writeln('  <v:textbox style="mso-direction-alt:auto"/>');
      buffer.writeln('  <x:ClientData ObjectType="Note">');
      buffer.writeln('   <x:MoveWithCells/>');
      buffer.writeln('   <x:SizeWithCells/>');
      buffer.writeln('   <x:Anchor>$anchor</x:Anchor>');
      buffer.writeln('   <x:AutoFill>False</x:AutoFill>');
      buffer.writeln('   <x:Row>$row</x:Row>');
      buffer.writeln('   <x:Column>$col</x:Column>');
      buffer.writeln('  </x:ClientData>');
      buffer.writeln(' </v:shape>');
    }

    buffer.writeln('</xml>');
    return buffer.toString();
  }
}
