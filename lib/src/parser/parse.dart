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
              XmlName.parts('Relationship'),
              <XmlAttribute>[
                XmlAttribute(XmlName.parts('Id'), 'rId$rIdNumber'),
                XmlAttribute(XmlName.parts('Type'),
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),
                XmlAttribute(XmlName.parts('Target'), 'sharedStrings.xml'),
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
                XmlName.parts('Override'),
                <XmlAttribute>[
                  XmlAttribute(
                      XmlName.parts('PartName'), '/xl/sharedStrings.xml'),
                  XmlAttribute(XmlName.parts('ContentType'), contentType),
                ],
              ));
        }
      }

      final emptyXml = utf8.encode(
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>');
      _excel._archive.addFile(
          ArchiveFile('xl/sharedStrings.xml', emptyXml.length, emptyXml));
      sharedStrings = _excel._archive.findFile('xl/sharedStrings.xml');
    }

    sharedStrings!.decompress();
    final contentString = utf8.decode(sharedStrings.content);
    final events = xml_events.parseEvents(contentString);
    List<xml_events.XmlEvent>? currentSiEvents;
    for (final event in events) {
      if (event is xml_events.XmlStartElementEvent &&
          (event.name == 'si' || event.name.endsWith(':si'))) {
        currentSiEvents = [event];
        if (event.isSelfClosing) {
          final siXml = event.toString();
          final node = XmlDocument.parse(siXml).rootElement;
          _parseSharedString(node);
          currentSiEvents = null;
        }
      } else if (currentSiEvents != null) {
        currentSiEvents.add(event);
        if (event is xml_events.XmlEndElementEvent &&
            (event.name == 'si' || event.name.endsWith(':si'))) {
          final siXml = currentSiEvents.map((e) => e.toString()).join();
          final node = XmlDocument.parse(siXml).rootElement;
          _parseSharedString(node);
          currentSiEvents = null;
        }
      }
    }
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
      final name = node.getAttribute('name');
      if (name != null) {
        _excel._sheets[name] = node;
      }
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
        .add(XmlElement(XmlName.parts('Relationship'), <XmlAttribute>[
          XmlAttribute(XmlName.parts('Id'), 'rId$ridNumber'),
          XmlAttribute(XmlName.parts('Type'), '$_relationships/worksheet'),
          XmlAttribute(
              XmlName.parts('Target'), 'worksheets/sheet$sheetNumber.xml'),
        ]));

    if (!_rId.contains('rId$ridNumber')) _rId.add('rId$ridNumber');

    final newSheetNode = XmlElement(XmlName.parts('sheet'), <XmlAttribute>[
      XmlAttribute(XmlName.parts('state'), 'visible'),
      XmlAttribute(XmlName.parts('name'), newSheet),
      XmlAttribute(XmlName.parts('sheetId'), '$sheetNumber'),
      XmlAttribute(XmlName.parts('r:id'), 'rId$ridNumber'),
    ]);

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheets')
        .first
        .children
        .add(newSheetNode);

    _excel._sheets[newSheet] = newSheetNode;

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
    final contentString = utf8.decode(newFile.content);
    _excel._sheetXmls['xl/worksheets/sheet$sheetNumber.xml'] = contentString;
    _excel._xmlSheetId[newSheet] = 'xl/worksheets/sheet$sheetNumber.xml';

    _excel._xmlFiles['[Content_Types].xml']
        ?.findAllElements('Types')
        .first
        .children
        .add(XmlElement(XmlName.parts('Override'), <XmlAttribute>[
          XmlAttribute(XmlName.parts('ContentType'),
              'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
          XmlAttribute(XmlName.parts('PartName'),
              '/xl/worksheets/sheet$sheetNumber.xml'),
        ]));

    if (_excel._xmlFiles['xl/workbook.xml'] != null) {
      _WorksheetParser(_excel, _worksheetTargets).parseTable(
          _excel._xmlFiles['xl/workbook.xml']!.findAllElements('sheet').last);
    }
  }
}
