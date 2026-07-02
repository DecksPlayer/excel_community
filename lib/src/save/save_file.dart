part of '../../excel_community.dart';

class Save {
  final Excel _excel;
  final Map<String, ArchiveFile> _archiveFiles = {};

  /// Binary (non-XML) files to include in the archive, e.g. images.
  final Map<String, List<int>> _binaryFiles = {};
  final List<CellStyle> _innerCellStyle = [];
  final Parser parser;
  late final _ChartManager _chartManager;
  late final _ImageManager _imageManager;
  late final _StyleManager _styleManager;
  late final _WorksheetManager _worksheetManager;
  late final _WorkbookManager _workbookManager;

  Save._(this._excel, this.parser) {
    _chartManager = _ChartManager(_excel, this);
    _imageManager = _ImageManager(_excel, this);
    _styleManager = _StyleManager(_excel, this);
    _worksheetManager = _WorksheetManager(_excel, this);
    _workbookManager = _WorkbookManager(_excel);
  }

  List<int>? _save() {
    _excel._sheetMap.forEach((sheetName, sheetObject) {
      if (!_excel._xmlSheetId.containsKey(sheetName)) {
        parser._createSheet(sheetName);
      }
    });

    if (_excel._styleChanges) {
      _styleManager.processStylesFile();
    }

    // Clear stale drawing references from the template for sheets that have no
    // charts and no images. The default _newSheet template already has a
    // drawing1.xml and an xl/worksheets/_rels/sheet1.xml.rels pointing to it,
    // which would cause unrelated sheets to incorrectly reference drawings
    // meant for other sheets.
    _excel._sheetMap.forEach((_, sheetObject) {
      if (sheetObject.charts.isEmpty && sheetObject.images.isEmpty) {
        sheetObject._drawingRId = null;
      }
    });

    _chartManager.processCharts();
    _imageManager.processImages();

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

    return ZipEncoder().encode(_cloneArchive(_excel._archive, _archiveFiles));
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
