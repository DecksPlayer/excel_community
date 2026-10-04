part of '../../../excel_community.dart';

/// Manages embedding images into the XLSX archive during save.
///
/// Runs after [_ChartManager] so it can detect existing drawing XML
/// (created by charts) and add image anchors to the same drawing,
/// since a worksheet can reference only one drawing part.
class _ImageManager {
  final Excel _excel;
  final Save _save;

  _ImageManager(this._excel, this._save);

  // -----------------------------------------------------------------------
  // PUBLIC
  // -----------------------------------------------------------------------

  void processImages() {
    // Continue after the media of the original file so it is not
    // overwritten.
    int imageCount = _save._highestPartIndex(
        RegExp(r'^xl/media/image(\d+)\.\w+$'),
        originalOnly: true);

    _excel._sheetMap.forEach((sheetName, sheet) {
      if (sheet.images.isEmpty) return;

      final sheetId = _excel._xmlSheetId[sheetName];
      if (sheetId == null) return;

      final sheetFileName = sheetId.split('/').last; // e.g. sheet1.xml
      final sheetRelsPath = 'xl/worksheets/_rels/$sheetFileName.rels';

      // --- Resolve or create the drawing for this sheet ---
      String drawingPath;
      String drawingRelsPath;
      bool drawingIsNew;

      final existing = _findExistingDrawing(sheetRelsPath);
      if (existing != null) {
        drawingPath = existing.$1;
        drawingRelsPath = existing.$2;
        drawingIsNew = false;
      } else {
        final idx =
            _save._highestPartIndex(RegExp(r'^xl/drawings/drawing(\d+)\.xml$')) + 1;
        drawingPath = 'xl/drawings/drawing$idx.xml';
        drawingRelsPath = 'xl/drawings/_rels/drawing$idx.xml.rels';
        drawingIsNew = true;
      }

      // --- Ensure drawing rels XML exists ---
      final drawingRels = _save._relationshipsPart(drawingRelsPath);
      final relsRoot = drawingRels.findAllElements('Relationships').first;
      int nextRId =
          int.parse(_save._nextRelationshipId(relsRoot).substring(3));

      // --- Ensure drawing XML exists ---
      var drawingDoc = _save._loadXmlPart(drawingPath);
      if (drawingDoc == null) {
        drawingDoc = _buildEmptyDrawing();
        _excel._xmlFiles[drawingPath] = drawingDoc;
      }
      // Use rootElement directly — the root of a drawing is always wsDr,
      // and findAllElements('wsDr') would fail because the xml package
      // matches qualified names ('xdr:wsDr' ≠ 'wsDr').
      final drawingRoot = drawingDoc.rootElement;

      // --- Process each image ---
      for (final image in sheet.images) {
        imageCount++;
        final mediaPath = 'xl/media/image$imageCount.${image._extension}';
        final rId = 'rId$nextRId';
        nextRId++;

        // Store binary bytes (written separately from XML files)
        _save._addMediaFile(mediaPath, List<int>.from(image.imageBytes));

        // Add image relationship in the drawing rels
        relsRoot.children.add(XmlElement(XmlName.parts('Relationship'), [
          XmlAttribute(XmlName.parts('Id'), rId),
          XmlAttribute(
            XmlName.parts('Type'),
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          ),
          XmlAttribute(
            XmlName.parts('Target'),
            '../media/image$imageCount.${image._extension}',
          ),
        ]));

        // Add <xdr:oneCellAnchor> to drawing XML
        drawingRoot.children
            .add(_buildImageAnchorElement(image, rId, imageCount));

        // Register content type for this image format (Default extension)
        _save._addDefaultContentType(image._contentType, image._extension);
      }

      // --- If drawing is new, wire it up to the worksheet ---
      if (drawingIsNew) {
        _save._addContentType(
          'application/vnd.openxmlformats-officedocument.drawing+xml',
          '/$drawingPath',
        );

        // Add drawing relationship to the sheet's rels file
        final sheetRels = _save._relationshipsPart(sheetRelsPath);
        final sheetRelsRoot = sheetRels.findAllElements('Relationships').first;
        final drawingRId = _save._nextRelationshipId(sheetRelsRoot);
        final drawingFileName = drawingPath.split('/').last;

        sheetRelsRoot.children.add(XmlElement(XmlName.parts('Relationship'), [
          XmlAttribute(XmlName.parts('Id'), drawingRId),
          XmlAttribute(
            XmlName.parts('Type'),
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          ),
          XmlAttribute(XmlName.parts('Target'), '../drawings/$drawingFileName'),
        ]));

        sheet._drawingRId = drawingRId;
      }
    });
  }

  // -----------------------------------------------------------------------
  // PRIVATE HELPERS
  // -----------------------------------------------------------------------

  /// Looks for an existing drawing relationship in the sheet's rels file.
  /// Returns (drawingPath, drawingRelsPath) or null.
  (String, String)? _findExistingDrawing(String sheetRelsPath) {
    final sheetRels = _save._loadXmlPart(sheetRelsPath);
    if (sheetRels == null) return null;

    for (final rel in sheetRels.findAllElements('Relationship')) {
      final type = rel.getAttribute('Type') ?? '';
      if (type.endsWith('/drawing')) {
        final target = rel.getAttribute('Target') ?? '';
        // target is like '../drawings/drawing1.xml'
        final drawingFileName = target.split('/').last;
        return (
          'xl/drawings/$drawingFileName',
          'xl/drawings/_rels/$drawingFileName.rels',
        );
      }
    }
    return null;
  }

  /// Counts drawing XML files already present in [_xmlFiles].


  XmlDocument _buildEmptyDrawing() {
    final b = XmlBuilder();
    b.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    b.element('xdr:wsDr',
        namespaceUris: {
          'xdr':
              'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
          'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'r':
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        },
        nest: () {});
    return b.buildDocument();
  }

  /// Builds a `<xdr:oneCellAnchor>` element for the given [image] using
  /// [XmlBuilder], keeping namespace handling consistent with [ChartXmlWriter].
  XmlElement _buildImageAnchorElement(
      ExcelImage image, String rId, int imageIndex) {
    final a = image.anchor;
    final b = XmlBuilder();
    b.element('xdr:oneCellAnchor', namespaceUris: {
      'xdr':
          'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }, nest: () {
      b.element('xdr:from', nest: () {
        b.element('xdr:col', nest: () => b.text(a.fromColumn.toString()));
        b.element('xdr:colOff', nest: () => b.text(a.colOffset.toString()));
        b.element('xdr:row', nest: () => b.text(a.fromRow.toString()));
        b.element('xdr:rowOff', nest: () => b.text(a.rowOffset.toString()));
      });
      b.element('xdr:ext', attributes: {
        'cx': a.widthEmu.toString(),
        'cy': a.heightEmu.toString(),
      });
      b.element('xdr:pic', nest: () {
        b.element('xdr:nvPicPr', nest: () {
          b.element('xdr:cNvPr', attributes: {
            'id': (imageIndex + 1).toString(),
            'name': 'Image $imageIndex',
          });
          b.element('xdr:cNvPicPr', nest: () {
            b.element('a:picLocks', attributes: {'noChangeAspect': '1'});
          });
        });
        b.element('xdr:blipFill', nest: () {
          b.element('a:blip', attributes: {'r:embed': rId});
          b.element('a:stretch', nest: () {
            b.element('a:fillRect');
          });
        });
        b.element('xdr:spPr', nest: () {
          b.element('a:xfrm', nest: () {
            b.element('a:off', attributes: {'x': '0', 'y': '0'});
            b.element('a:ext', attributes: {
              'cx': a.widthEmu.toString(),
              'cy': a.heightEmu.toString(),
            });
          });
          b.element('a:prstGeom', attributes: {'prst': 'rect'}, nest: () {
            b.element('a:avLst');
          });
        });
      });
      b.element('xdr:clientData');
    });
    return b.buildDocument().rootElement.copy();
  }
}
