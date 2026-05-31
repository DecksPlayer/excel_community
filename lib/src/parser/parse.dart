part of '../../excel_community.dart';

/// Orchestrates the full parsing of an `.xlsx` archive.
///
/// Responsibilities kept here:
///   - ZIP / content-type bootstrap ([_putContentXml])
///   - Relationship resolution ([_parseRelations])
///   - Shared strings ([_parseSharedStrings])
///   - Workbook content iteration ([_parseContent])
///   - Merged-cell processing ([_parseMergedCells])
///   - New-sheet creation ([_createSheet])
///
/// Styles parsing is delegated to [_StylesParser].
/// Worksheet / row / cell parsing is delegated to [_WorksheetParser].
class Parser {
  final Excel _excel;

  /// Relationship IDs found in the workbook relations file.
  final List<String> _rId = [];

  /// Maps each relationship ID to its worksheet target path.
  final Map<String, String> _worksheetTargets = {};

  Parser._(this._excel);

  // ---------------------------------------------------------------------------
  // Entry point
  // ---------------------------------------------------------------------------

  void _startParsing() {
    _putContentXml();
    _parseRelations();
    _StylesParser(_excel).parse(_excel._stylesTarget);
    _parseSharedStrings();
    _parseContent();
    _parseMergedCells();
  }

  // ---------------------------------------------------------------------------
  // Bootstrap
  // ---------------------------------------------------------------------------

  void _putContentXml() {
    var file = _excel._archive.findFile('[Content_Types].xml');
    if (file == null) {
      _damagedExcel();
    }
    file!.decompress();
    _excel._xmlFiles['[Content_Types].xml'] =
        XmlDocument.parse(utf8.decode(file.content));
  }

  void _parseRelations() {
    var relations = _excel._archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations == null) {
      _damagedExcel();
      return;
    }

    relations.decompress();
    var document = XmlDocument.parse(utf8.decode(relations.content));
    _excel._xmlFiles['xl/_rels/workbook.xml.rels'] = document;

    document.findAllElements('Relationship').forEach((node) {
      final id = node.getAttribute('Id');
      final target = node.getAttribute('Target');
      if (target != null) {
        switch (node.getAttribute('Type')) {
          case _relationshipsStyles:
            _excel._stylesTarget = target;
            break;
          case _relationshipsWorksheet:
            if (id != null) _worksheetTargets[id] = target;
            break;
          case _relationshipsSharedStrings:
            _excel._sharedStringsTarget = target;
            break;
        }
      }
      if (id != null && !_rId.contains(id)) {
        _rId.add(id);
      }
    });
  }

  // ---------------------------------------------------------------------------
  // Shared strings
  // ---------------------------------------------------------------------------

  void _parseSharedStrings() {
    var sharedStrings =
        _excel._archive.findFile(_excel._absSharedStringsTarget);

    if (sharedStrings == null) {
      // File doesn't exist yet — create an empty one and wire it up.
      _excel._sharedStringsTarget = 'sharedStrings.xml';

      // Run without parsing to collect all rIds first.
      _parseContent(run: false);

      if (_excel._xmlFiles.containsKey('xl/_rels/workbook.xml.rels')) {
        final rIdNumber = _getAvailableRid();

        _excel._xmlFiles['xl/_rels/workbook.xml.rels']
            ?.findAllElements('Relationships')
            .first
            .children
            .add(XmlElement(
              XmlName('Relationship'),
              <XmlAttribute>[
                XmlAttribute(XmlName('Id'), 'rId$rIdNumber'),
                XmlAttribute(XmlName('Type'),
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),
                XmlAttribute(XmlName('Target'), 'sharedStrings.xml'),
              ],
            ));

        if (!_rId.contains('rId$rIdNumber')) _rId.add('rId$rIdNumber');

        const contentType =
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
        var alreadyPresent = false;

        _excel._xmlFiles['[Content_Types].xml']
            ?.findAllElements('Override')
            .forEach((node) {
          if (node.getAttribute('ContentType') == contentType) {
            alreadyPresent = true;
          }
        });

        if (!alreadyPresent) {
          _excel._xmlFiles['[Content_Types].xml']
              ?.findAllElements('Types')
              .first
              .children
              .add(XmlElement(
                XmlName('Override'),
                <XmlAttribute>[
                  XmlAttribute(XmlName('PartName'), '/xl/sharedStrings.xml'),
                  XmlAttribute(XmlName('ContentType'), contentType),
                ],
              ));
        }
      }

      final emptyXml = utf8.encode(
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>');
      _excel._archive
          .addFile(ArchiveFile('xl/sharedStrings.xml', emptyXml.length, emptyXml));
      sharedStrings = _excel._archive.findFile('xl/sharedStrings.xml');
    }

    sharedStrings!.decompress();
    var document = XmlDocument.parse(utf8.decode(sharedStrings.content));
    _excel._xmlFiles['xl/${_excel._sharedStringsTarget}'] = document;

    document.findAllElements('si').forEach(_parseSharedString);
  }

  void _parseSharedString(XmlElement node) {
    final sharedString = SharedString(node: node);
    _excel._sharedStrings.add(sharedString, sharedString.stringValue);
  }

  // ---------------------------------------------------------------------------
  // Workbook content
  // ---------------------------------------------------------------------------

  void _parseContent({bool run = true}) {
    var workbook = _excel._archive.findFile('xl/workbook.xml');
    if (workbook == null) {
      _damagedExcel();
      return;
    }
    workbook.decompress();
    var document = XmlDocument.parse(utf8.decode(workbook.content));
    _excel._xmlFiles['xl/workbook.xml'] = document;

    document.findAllElements('sheet').forEach((node) {
      if (run) {
        _WorksheetParser(_excel, _worksheetTargets).parseTable(node);
      } else {
        final rid = node.getAttribute('r:id');
        if (rid != null && !_rId.contains(rid)) {
          _rId.add(rid);
        }
      }
    });
  }

  // ---------------------------------------------------------------------------
  // Merged cells
  // ---------------------------------------------------------------------------

  /// Identifies merged cell regions in each sheet and removes all spanned
  /// cells except the top-left one, which preserves its content.
  void _parseMergedCells() {
    final spannedCells = <String, List<String>>{};

    _excel._sheets.forEach((sheetName, node) {
      _excel._availSheet(sheetName);
      final sheetDataNode = node as XmlElement;
      final spanList = <String>[];
      final sheet = _excel._sheetMap[sheetName]!;

      sheetDataNode.parent!.findAllElements('mergeCell').forEach((element) {
        final ref = element.getAttribute('ref');
        if (ref == null || !ref.contains(':') || ref.split(':').length != 2) {
          return;
        }

        if (!sheet._spannedItems.contains(ref)) {
          sheet._spannedItems.add(ref);
        }

        final parts = ref.split(':');
        final startCell = parts[0];
        final endCell = parts[1];

        if (!spanList.contains(startCell)) spanList.add(startCell);
        spannedCells[sheetName] = spanList;

        final spanObj = _Span.fromCellIndex(
          start: CellIndex.indexByString(startCell),
          end: CellIndex.indexByString(endCell),
        );

        if (!sheet._spanList.contains(spanObj)) {
          sheet._spanList.add(spanObj);
          _clearMergedCells(spanObj, sheet);
        }
        _excel._mergeChangeLookup = sheetName;
      });
    });
  }

  /// Removes all cells inside a merged span except the top-left cell.
  void _clearMergedCells(_Span spanObj, Sheet sheet) {
    for (var col = spanObj.columnSpanStart; col <= spanObj.columnSpanEnd; col++) {
      for (var row = spanObj.rowSpanStart; row <= spanObj.rowSpanEnd; row++) {
        final isOrigin = col == spanObj.columnSpanStart && row == spanObj.rowSpanStart;
        if (!isOrigin) sheet._removeCell(row, col);
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Sheet creation
  // ---------------------------------------------------------------------------

  int _getAvailableRid() {
    _rId.sort((a, b) =>
        int.parse(a.substring(3)).compareTo(int.parse(b.substring(3))));

    final digits = _rId.last.replaceAll(RegExp(r'[^0-9]'), '');
    return int.parse(digits) + 1;
  }

  /// Creates a new empty worksheet, wires up all relationships and XML files,
  /// then parses the newly created sheet so it is immediately usable.
  void _createSheet(String newSheet) {
    int sheetId = -1;
    final sheetIdList = <int>[];

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheet')
        .forEach((node) {
      final raw = node.getAttribute('sheetId');
      if (raw != null) {
        final id = int.parse(raw);
        if (!sheetIdList.contains(id)) sheetIdList.add(id);
      } else {
        _damagedExcel(text: 'Corrupted Sheet Indexing');
      }
    });

    sheetIdList.sort();
    for (int i = 0; i < sheetIdList.length; i++) {
      if ((i + 1) != sheetIdList[i]) {
        sheetId = i + 1;
        break;
      }
    }
    if (sheetId == -1) {
      sheetId = sheetIdList.isEmpty ? 1 : sheetIdList.length + 1;
    }

    final sheetNumber = sheetId;
    final ridNumber = _getAvailableRid();

    _excel._xmlFiles['xl/_rels/workbook.xml.rels']
        ?.findAllElements('Relationships')
        .first
        .children
        .add(XmlElement(XmlName('Relationship'), <XmlAttribute>[
          XmlAttribute(XmlName('Id'), 'rId$ridNumber'),
          XmlAttribute(XmlName('Type'), '$_relationships/worksheet'),
          XmlAttribute(
              XmlName('Target'), 'worksheets/sheet$sheetNumber.xml'),
        ]));

    if (!_rId.contains('rId$ridNumber')) _rId.add('rId$ridNumber');

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheets')
        .first
        .children
        .add(XmlElement(XmlName('sheet'), <XmlAttribute>[
          XmlAttribute(XmlName('state'), 'visible'),
          XmlAttribute(XmlName('name'), newSheet),
          XmlAttribute(XmlName('sheetId'), '$sheetNumber'),
          XmlAttribute(XmlName('r:id'), 'rId$ridNumber'),
        ]));

    _worksheetTargets['rId$ridNumber'] = 'worksheets/sheet$sheetNumber.xml';

    final blankXml = utf8.encode(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' mc:Ignorable="x14ac xr xr2 xr3"'
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
        ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"'
        ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"'
        ' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">'
        ' <dimension ref="A1"/>'
        ' <sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        ' <sheetData/>'
        ' <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        ' </worksheet>');

    _excel._archive.addFile(ArchiveFile(
        'xl/worksheets/sheet$sheetNumber.xml', blankXml.length, blankXml));

    final newFile =
        _excel._archive.findFile('xl/worksheets/sheet$sheetNumber.xml')!;
    newFile.decompress();
    final document = XmlDocument.parse(utf8.decode(newFile.content));
    _excel._xmlFiles['xl/worksheets/sheet$sheetNumber.xml'] = document;
    _excel._xmlSheetId[newSheet] = 'xl/worksheets/sheet$sheetNumber.xml';

    _excel._xmlFiles['[Content_Types].xml']
        ?.findAllElements('Types')
        .first
        .children
        .add(XmlElement(XmlName('Override'), <XmlAttribute>[
          XmlAttribute(XmlName('ContentType'),
              'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
          XmlAttribute(
              XmlName('PartName'), '/xl/worksheets/sheet$sheetNumber.xml'),
        ]));

    if (_excel._xmlFiles['xl/workbook.xml'] != null) {
      _WorksheetParser(_excel, _worksheetTargets).parseTable(
          _excel._xmlFiles['xl/workbook.xml']!
              .findAllElements('sheet')
              .last);
    }
  }
}
