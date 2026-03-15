part of excel_community;

class Save {
  final Excel _excel;
  final Map<String, ArchiveFile> _archiveFiles = {};
  final List<CellStyle> _innerCellStyle = [];
  final Parser parser;
  late final _ChartManager _chartManager;
  late final _StyleManager _styleManager;
  late final _WorksheetManager _worksheetManager;
  late final _WorkbookManager _workbookManager;

  Save._(this._excel, this.parser) {
    _chartManager = _ChartManager(_excel, this);
    _styleManager = _StyleManager(_excel, this);
    _worksheetManager = _WorksheetManager(_excel, this);
    _workbookManager = _WorkbookManager(_excel, this);
  }

  List<int>? _save() {
    if (_excel._styleChanges) {
      _styleManager.processStylesFile();
    }
    
    _worksheetManager.setSheetElements();

    if (_excel._defaultSheet != null) {
      _workbookManager.setDefaultSheet(_excel._defaultSheet);
    }

    _workbookManager.setSharedStrings();

    if (_excel._mergeChanges) {
      _workbookManager.setMerge();
    }

    if (_excel._rtlChanges) {
      _workbookManager.setRTL();
    }

    _chartManager.processCharts();

    for (var xmlFile in _excel._xmlFiles.keys) {
      var xml = _excel._xmlFiles[xmlFile].toString();
      var content = utf8.encode(xml);
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
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

  void _setHeaderFooter(String sheetName) {
    _workbookManager.setHeaderFooter(sheetName);
  }

  void _addContentType(String contentType, String partName) {
    final contentTypes = _excel._xmlFiles['[Content_Types].xml'];
    if (contentTypes == null) return;

    final typesElement = contentTypes.findAllElements('Types').first;

    // Check if already exists
    final exists = typesElement.children.any((node) =>
        node is XmlElement && node.getAttribute('PartName') == partName);

    if (!exists) {
      typesElement.children.add(XmlElement(XmlName('Override'), [
        XmlAttribute(XmlName('PartName'), partName),
        XmlAttribute(XmlName('ContentType'), contentType),
      ]));
    }
  }
}
