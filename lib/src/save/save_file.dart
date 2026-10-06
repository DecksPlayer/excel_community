part of '../../excel_community.dart';

class Save {
  final Excel _excel;
  final Map<String, ArchiveFile> _archiveFiles = {};

  /// Binary (non-XML) files to include in the archive, e.g. images.
  final Map<String, List<int>> _binaryFiles = {};
  final List<CellStyle> _innerCellStyle = [];

  /// Parts of the original file left out of the saved archive.
  final Set<String> _excludedParts = {};
  final Parser parser;
  late final _ChartManager _chartManager;
  late final _ImageManager _imageManager;
  late final _PivotTableManager _pivotTableManager;
  late final _StyleManager _styleManager;
  late final _WorksheetManager _worksheetManager;
  late final _WorkbookManager _workbookManager;
  late final _CommentManager _commentManager;
  late final _HyperlinkManager _hyperlinkManager;
  late final _TableManager _tableManager;

  Save._(this._excel, this.parser) {
    _chartManager = _ChartManager(_excel, this);
    _imageManager = _ImageManager(_excel, this);
    _pivotTableManager = _PivotTableManager(_excel, this);
    _styleManager = _StyleManager(_excel, this);
    _worksheetManager = _WorksheetManager(_excel, this);
    _workbookManager = _WorkbookManager(_excel);
    _commentManager = _CommentManager(_excel, this);
    _hyperlinkManager = _HyperlinkManager(_excel, this);
    _tableManager = _TableManager(_excel, this);
  }

  List<int>? _save() {
    _excel._sheetMap.forEach((sheetName, sheetObject) {
      if (!_excel._xmlSheetId.containsKey(sheetName)) {
        parser._createSheet(sheetName);
      }
    });

    // Saving appends drawings, relationships, styles and content types to the
    // workbook's XML parts. Work on copies so the workbook itself is left
    // unchanged: otherwise every save appended them again, duplicating chart
    // and image anchors, and Excel had to repair the file.
    final xmlFiles = Map.of(_excel._xmlFiles);
    final sheetXmls = Map.of(_excel._sheetXmls);
    _excel._xmlFiles.updateAll((_, document) => document.copy());
    try {
      return _writeArchive();
    } finally {
      _excel._xmlFiles
        ..clear()
        ..addAll(xmlFiles);
      _excel._sheetXmls
        ..clear()
        ..addAll(sheetXmls);
    }
  }

  List<int>? _writeArchive() {
    _dropCalcChain();

    // Write table header/totals cells and pivot table results, so they run
    // before styles.
    _tableManager.syncTables();
    _pivotTableManager.renderPivotTables();

    if (_excel._styleChanges) {
      _styleManager.processStylesFile();
    }

    _chartManager.processCharts();
    _imageManager.processImages();
    _pivotTableManager.processPivotTables();
    _commentManager.processComments();
    _hyperlinkManager.processHyperlinks();
    _tableManager.processTables();

    _worksheetManager.setSheetElements();

    if (_excel._defaultSheet != null) {
      _workbookManager.setDefaultSheet(_excel._defaultSheet);
    }

    final sstXml = _workbookManager.generateSharedStringsXml();
    final sstBytes = utf8.encode(sstXml);
    _archiveFiles[_excel._absSharedStringsTarget] = ArchiveFile(
      _excel._absSharedStringsTarget,
      sstBytes.length,
      sstBytes,
    );

    for (var xmlFile in _excel._xmlFiles.keys) {
      if (xmlFile == 'xl/${_excel._sharedStringsTarget}' ||
          xmlFile == _excel._absSharedStringsTarget) {
        continue;
      }
      var xml = _excel._xmlFiles[xmlFile].toString();
      var content = utf8.encode(xml);
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
    }

    for (var sheetPath in _excel._sheetXmls.keys) {
      var xml = _excel._sheetXmls[sheetPath]!;
      var content = utf8.encode(xml);
      _archiveFiles[sheetPath] =
          ArchiveFile(sheetPath, content.length, content);
    }

    for (final entry in _binaryFiles.entries) {
      _archiveFiles[entry.key] =
          ArchiveFile(entry.key, entry.value.length, entry.value);
    }

    return ZipEncoder().encode(_cloneArchive(_excel._archive, _archiveFiles,
        excludedFiles: _excludedParts));
  }

  _BorderSet _createBorderSetFromCellStyle(CellStyle cellStyle) => _BorderSet(
        leftBorder: cellStyle.leftBorder,
        rightBorder: cellStyle.rightBorder,
        topBorder: cellStyle.topBorder,
        bottomBorder: cellStyle.bottomBorder,
        diagonalBorder: cellStyle.diagonalBorder,
        diagonalBorderUp: cellStyle.diagonalBorderUp,
        diagonalBorderDown: cellStyle.diagonalBorderDown,
      );

  /// Leaves out Excel's calculation chain (`xl/calcChain.xml`).
  ///
  /// It lists formula cells by position, so inserting or removing rows,
  /// moving tables or replacing formulas makes it stale, and Excel then
  /// reports the file as damaged. Excel rebuilds it when it is missing.
  void _dropCalcChain() {
    const relType =
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';
    final rels = _excel._xmlFiles['xl/_rels/workbook.xml.rels'];
    final removed = <XmlElement>[];
    for (final rel in rels?.findAllElements('Relationship') ?? const <XmlElement>[]) {
      if (rel.getAttribute('Type') != relType) continue;
      removed.add(rel);
      final target = rel.getAttribute('Target') ?? 'calcChain.xml';
      final path = target.startsWith('/') ? target.substring(1) : 'xl/$target';
      _excludedParts.add(path);
      _excel._xmlFiles.remove(path);
      _excel._xmlFiles['[Content_Types].xml']?.rootElement.children.removeWhere((node) =>
          node is XmlElement && node.getAttribute('PartName') == '/$path');
    }
    for (final rel in removed) {
      rel.parent?.children.remove(rel);
    }
  }

  /// Returns the XML part at [path]: the copy already loaded or created
  /// during this save, otherwise the part parsed from the original file
  /// (cached in `_xmlFiles` so later edits are written back), or `null`.
  XmlDocument? _loadXmlPart(String path) {
    final loaded = _excel._xmlFiles[path];
    if (loaded != null) return loaded;
    final file = _excel._archive.findFile(path);
    if (file == null) return null;
    file.decompress();
    final document = XmlDocument.parse(utf8.decode(file.content));
    _excel._xmlFiles[path] = document;
    return document;
  }

  /// Like [_loadXmlPart] for a `.rels` part, creating an empty
  /// `<Relationships>` part when it does not exist yet.
  XmlDocument _relationshipsPart(String path) {
    final existing = _loadXmlPart(path);
    if (existing != null) return existing;
    final builder = XmlBuilder();
    builder.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('Relationships', attributes: {
      'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
    });
    final document = builder.buildDocument();
    _excel._xmlFiles[path] = document;
    return document;
  }

  /// Next unused `rIdN` in a `<Relationships>` element (highest + 1, so ids
  /// with gaps never collide).
  String _nextRelationshipId(XmlElement relationships) {
    var highest = 0;
    for (final rel in relationships.childElements) {
      final match = RegExp(r'^rId(\d+)$').firstMatch(rel.getAttribute('Id') ?? '');
      if (match != null) highest = max(highest, int.parse(match.group(1)!));
    }
    return 'rId${highest + 1}';
  }

  /// Highest N among part names matching [pattern] (whose first group is
  /// N), e.g. `^xl/drawings/drawing(\d+)\.xml$`. Looks at the original file
  /// and, unless [originalOnly], at the parts created during this save.
  int _highestPartIndex(RegExp pattern, {bool originalOnly = false}) {
    var highest = 0;
    final names = {
      ..._excel._archive.files.map((f) => f.name),
      if (!originalOnly) ..._excel._xmlFiles.keys,
    };
    for (final name in names) {
      final match = pattern.firstMatch(name);
      if (match == null) continue;
      highest = max(highest, int.parse(match.group(1)!));
    }
    return highest;
  }

  void _addContentType(String contentType, String partName) {
    final contentTypes = _excel._xmlFiles['[Content_Types].xml'];
    if (contentTypes == null) return;

    final typesElement = contentTypes.findAllElements('Types').first;

    // Check if already exists
    final exists = typesElement.children.any((node) =>
        node is XmlElement && node.getAttribute('PartName') == partName);

    if (!exists) {
      typesElement.children.add(XmlElement(XmlName.parts('Override'), [
        XmlAttribute(XmlName.parts('PartName'), partName),
        XmlAttribute(XmlName.parts('ContentType'), contentType),
      ]));
    }
  }

  /// Registers a `<Default Extension="…" ContentType="…"/>` entry in
  /// [Content_Types].xml (used for binary media such as images).
  void _addDefaultContentType(String contentType, String extension) {
    final contentTypes = _excel._xmlFiles['[Content_Types].xml'];
    if (contentTypes == null) return;

    final typesElement = contentTypes.findAllElements('Types').first;

    final exists = typesElement.children.any((node) =>
        node is XmlElement &&
        node.name.local == 'Default' &&
        node.getAttribute('Extension') == extension);

    if (!exists) {
      typesElement.children.add(XmlElement(XmlName.parts('Default'), [
        XmlAttribute(XmlName.parts('Extension'), extension),
        XmlAttribute(XmlName.parts('ContentType'), contentType),
      ]));
    }
  }

  /// Stores binary [bytes] at [path] inside the XLSX archive.
  void _addMediaFile(String path, List<int> bytes) {
    _binaryFiles[path] = bytes;
  }
}
