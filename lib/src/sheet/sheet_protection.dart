part of '../../excel_community.dart';

class SheetProtection {
  bool sheet;
  bool objects;
  bool scenarios;
  bool formatCells;
  bool formatColumns;
  bool formatRows;
  bool insertColumns;
  bool insertRows;
  bool insertHyperlinks;
  bool deleteColumns;
  bool deleteRows;
  bool selectLockedCells;
  bool selectUnlockedCells;
  bool sort;
  bool autoFilter;
  bool pivotTables;
  String? password; // Hashed password hex string

  SheetProtection({
    this.sheet = false,
    this.objects = true,
    this.scenarios = true,
    this.formatCells = true,
    this.formatColumns = true,
    this.formatRows = true,
    this.insertColumns = true,
    this.insertRows = true,
    this.insertHyperlinks = true,
    this.deleteColumns = true,
    this.deleteRows = true,
    this.selectLockedCells = true,
    this.selectUnlockedCells = true,
    this.sort = true,
    this.autoFilter = true,
    this.pivotTables = true,
    this.password,
  });

  /// Set raw password (computes the legacy 16-bit hash and sets it in uppercase hex format)
  void setPassword(String rawPassword) {
    password = _getPasswordHash(rawPassword);
  }

  String toXmlString() {
    final sb = StringBuffer('<sheetProtection');
    sb.write(' sheet="${sheet ? "1" : "0"}"');
    sb.write(' objects="${objects ? "1" : "0"}"');
    sb.write(' scenarios="${scenarios ? "1" : "0"}"');
    sb.write(' formatCells="${formatCells ? "1" : "0"}"');
    sb.write(' formatColumns="${formatColumns ? "1" : "0"}"');
    sb.write(' formatRows="${formatRows ? "1" : "0"}"');
    sb.write(' insertColumns="${insertColumns ? "1" : "0"}"');
    sb.write(' insertRows="${insertRows ? "1" : "0"}"');
    sb.write(' insertHyperlinks="${insertHyperlinks ? "1" : "0"}"');
    sb.write(' deleteColumns="${deleteColumns ? "1" : "0"}"');
    sb.write(' deleteRows="${deleteRows ? "1" : "0"}"');
    sb.write(' selectLockedCells="${selectLockedCells ? "1" : "0"}"');
    sb.write(' selectUnlockedCells="${selectUnlockedCells ? "1" : "0"}"');
    sb.write(' sort="${sort ? "1" : "0"}"');
    sb.write(' autoFilter="${autoFilter ? "1" : "0"}"');
    sb.write(' pivotTables="${pivotTables ? "1" : "0"}"');
    if (password != null && password!.isNotEmpty) {
      sb.write(' password="$password"');
    }
    sb.write('/>');
    return sb.toString();
  }
}

/// 16-bit legacy password hashing algorithm for worksheet protection
String _getPasswordHash(String password) {
  int hash = 0;
  if (password.isNotEmpty) {
    final charCodes = password.codeUnits;
    for (int i = charCodes.length - 1; i >= 0; i--) {
      final char = charCodes[i];
      hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
      hash = hash ^ char;
    }
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    hash = hash ^ password.length ^ 0xCE4B;
  }
  return hash.toRadixString(16).toUpperCase().padLeft(4, '0');
}
