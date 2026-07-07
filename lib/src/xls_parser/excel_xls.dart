import 'dart:typed_data';
import '../../excel_community.dart';

/// Entry point to decode XLS (Excel 97-2003) file bytes.
Excel decodeXlsBytes(List<int> bytes) {
  final cfb = CfbFile(bytes);
  final workbookStream = cfb.getStream("Workbook") ?? cfb.getStream("Book");
  if (workbookStream == null) {
    throw UnsupportedError("XLS file does not contain a Workbook stream.");
  }
  return BiffParser().parse(workbookStream);
}

class CfbFile {
  final List<int> bytes;
  final ByteData view;
  late int sectorSize;
  late int miniSectorSize;
  late List<int> fat;
  late List<int> miniFat;
  late List<int> miniStream;
  late List<CfbDirectoryEntry> directories;

  CfbFile(this.bytes) : view = ByteData.sublistView(Uint8List.fromList(bytes)) {
    _parseHeader();
  }

  void _parseHeader() {
    if (bytes.length < 512) {
      throw UnsupportedError("Invalid XLS file (too short).");
    }
    final magic = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    for (int i = 0; i < 8; i++) {
      if (bytes[i] != magic[i]) {
        throw UnsupportedError("Invalid XLS signature.");
      }
    }

    sectorSize = 1 << view.getUint16(30, Endian.little);
    miniSectorSize = 1 << view.getUint16(32, Endian.little);

    final dirStartSector = view.getUint32(48, Endian.little);
    final miniFatStartSector = view.getUint32(60, Endian.little);
    final numMiniFatSectors = view.getUint32(64, Endian.little);
    final difatStartSector = view.getUint32(68, Endian.little);
    final numDifatSectors = view.getUint32(72, Endian.little);

    // Read DIFAT
    List<int> fatSectors = [];
    for (int i = 0; i < 109; i++) {
      int sec = view.getUint32(76 + i * 4, Endian.little);
      if (sec < 0xFFFFFFFC) {
        fatSectors.add(sec);
      }
    }

    if (numDifatSectors > 0 && difatStartSector < 0xFFFFFFFC) {
      int currentDifatSector = difatStartSector;
      while (currentDifatSector >= 0 && currentDifatSector < 0xFFFFFFFC) {
        int offset = 512 + currentDifatSector * sectorSize;
        int numEntries = (sectorSize ~/ 4) - 1;
        for (int i = 0; i < numEntries; i++) {
          int sec = view.getUint32(offset + i * 4, Endian.little);
          if (sec < 0xFFFFFFFC) {
            fatSectors.add(sec);
          }
        }
        currentDifatSector = view.getUint32(offset + numEntries * 4, Endian.little);
      }
    }

    // Build standard FAT
    fat = [];
    for (int fatSec in fatSectors) {
      int offset = 512 + fatSec * sectorSize;
      int numInts = sectorSize ~/ 4;
      for (int i = 0; i < numInts; i++) {
        fat.add(view.getUint32(offset + i * 4, Endian.little));
      }
    }

    // Parse Directory Entries
    final dirBytes = _readSectorChain(dirStartSector, -1);
    directories = [];
    for (int i = 0; i < dirBytes.length; i += 128) {
      if (i + 128 > dirBytes.length) break;
      final entryBytes = dirBytes.sublist(i, i + 128);
      final entryView = ByteData.sublistView(Uint8List.fromList(entryBytes));
      
      int nameLen = entryView.getUint16(64, Endian.little);
      String name = "";
      if (nameLen > 2) {
        List<int> nameUtf16 = entryBytes.sublist(0, nameLen - 2);
        StringBuffer sb = StringBuffer();
        for (int j = 0; j < nameUtf16.length; j += 2) {
          if (j + 1 < nameUtf16.length) {
            sb.writeCharCode(nameUtf16[j] | (nameUtf16[j + 1] << 8));
          }
        }
        name = sb.toString();
      }

      int type = entryView.getUint8(66);
      int startSec = entryView.getUint32(116, Endian.little);
      int size = entryView.getUint32(120, Endian.little);
      directories.add(CfbDirectoryEntry(name, type, startSec, size));
    }

    // Parse Mini Stream and Mini FAT
    final root = directories.isNotEmpty ? directories[0] : null;
    if (root != null && root.startSector < 0xFFFFFFFC && root.streamSize > 0) {
      miniStream = _readSectorChain(root.startSector, root.streamSize);
    } else {
      miniStream = [];
    }

    if (miniFatStartSector < 0xFFFFFFFC && numMiniFatSectors > 0) {
      final miniFatBytes = _readSectorChain(miniFatStartSector, -1);
      final mfView = ByteData.sublistView(Uint8List.fromList(miniFatBytes));
      miniFat = [];
      for (int i = 0; i < miniFatBytes.length; i += 4) {
        miniFat.add(mfView.getUint32(i, Endian.little));
      }
    } else {
      miniFat = [];
    }
  }

  List<int> _readSectorChain(int startSector, int size) {
    List<int> chain = [];
    int currentSector = startSector;
    while (currentSector >= 0 && currentSector < 0xFFFFFFFC) {
      int offset = 512 + currentSector * sectorSize;
      if (offset + sectorSize > bytes.length) {
        chain.addAll(bytes.sublist(offset, bytes.length));
        break;
      }
      chain.addAll(bytes.sublist(offset, offset + sectorSize));
      if (currentSector >= fat.length) break;
      currentSector = fat[currentSector];
    }
    if (size >= 0 && chain.length > size) {
      return chain.sublist(0, size);
    }
    return chain;
  }

  List<int>? getStream(String name) {
    for (final dir in directories) {
      if (dir.name.toLowerCase() == name.toLowerCase()) {
        if (dir.streamSize < 4096) {
          List<int> chain = [];
          int currentSector = dir.startSector;
          while (currentSector >= 0 && currentSector < 0xFFFFFFFC) {
            int offset = currentSector * miniSectorSize;
            if (offset + miniSectorSize > miniStream.length) {
              chain.addAll(miniStream.sublist(offset, miniStream.length));
              break;
            }
            chain.addAll(miniStream.sublist(offset, offset + miniSectorSize));
            if (currentSector >= miniFat.length) break;
            currentSector = miniFat[currentSector];
          }
          if (chain.length > dir.streamSize) {
            return chain.sublist(0, dir.streamSize);
          }
          return chain;
        } else {
          return _readSectorChain(dir.startSector, dir.streamSize);
        }
      }
    }
    return null;
  }
}

class CfbDirectoryEntry {
  final String name;
  final int type;
  final int startSector;
  final int streamSize;
  CfbDirectoryEntry(this.name, this.type, this.startSector, this.streamSize);
}

class BiffParser {
  Excel parse(List<int> workbookStream) {
    final excel = Excel.createExcel();
    final records = _readRecords(workbookStream);

    List<String> sst = [];
    int sstRecordIndex = -1;
    for (int i = 0; i < records.length; i++) {
      if (records[i].type == 0x00FC) {
        sstRecordIndex = i;
        break;
      }
    }

    if (sstRecordIndex != -1) {
      sst = _parseSst(records, sstRecordIndex);
    }

    final sheets = <_BiffSheet>[];
    for (final rec in records) {
      if (rec.type == 0x0085) {
        final view = ByteData.sublistView(Uint8List.fromList(rec.data));
        int offset = view.getUint32(0, Endian.little);
        int sheetType = view.getUint8(4);
        int charCount = view.getUint8(6);
        int flags = view.getUint8(7);
        
        bool isUtf16 = (flags & 0x01) != 0;
        String name = "";
        List<int> nameBytes = rec.data.sublist(8);
        if (isUtf16) {
          StringBuffer sb = StringBuffer();
          for (int j = 0; j < charCount * 2; j += 2) {
            if (j + 1 < nameBytes.length) {
              sb.writeCharCode(nameBytes[j] | (nameBytes[j + 1] << 8));
            }
          }
          name = sb.toString();
        } else {
          name = String.fromCharCodes(nameBytes.sublist(0, charCount));
        }

        if (sheetType == 0x00) {
          sheets.add(_BiffSheet(name, offset));
        }
      }
    }

    bool hasSheet1 = false;
    for (final sheetInfo in sheets) {
      final name = sheetInfo.name;
      if (name == 'Sheet1') {
        hasSheet1 = true;
      }
      final sheet = excel[name];

      int sheetStartIndex = -1;
      for (int i = 0; i < records.length; i++) {
        if (records[i].type == 0x0809 && records[i].streamOffset == sheetInfo.offset) {
          sheetStartIndex = i;
          break;
        }
      }

      if (sheetStartIndex != -1) {
        for (int i = sheetStartIndex + 1; i < records.length; i++) {
          final rec = records[i];
          if (rec.type == 0x000A) {
            break;
          }

          if (rec.type == 0x00FD) { // LABELSST
            final view = ByteData.sublistView(Uint8List.fromList(rec.data));
            int r = view.getUint16(0, Endian.little);
            int c = view.getUint16(2, Endian.little);
            int sstIndex = view.getUint32(6, Endian.little);
            if (sstIndex < sst.length) {
              sheet.updateCell(
                CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
                TextCellValue(sst[sstIndex]),
              );
            }
          } else if (rec.type == 0x0203) { // NUMBER
            final view = ByteData.sublistView(Uint8List.fromList(rec.data));
            int r = view.getUint16(0, Endian.little);
            int c = view.getUint16(2, Endian.little);
            double val = view.getFloat64(6, Endian.little);
            sheet.updateCell(
              CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
              DoubleCellValue(val),
            );
          } else if (rec.type == 0x027E) { // RK
            final view = ByteData.sublistView(Uint8List.fromList(rec.data));
            int r = view.getUint16(0, Endian.little);
            int c = view.getUint16(2, Endian.little);
            int rkValue = view.getUint32(6, Endian.little);
            final val = _decodeRk(rkValue);
            CellValue? cellValue;
            if (val is int) {
              cellValue = IntCellValue(val);
            } else if (val is double) {
              cellValue = DoubleCellValue(val);
            }
            if (cellValue != null) {
              sheet.updateCell(
                CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
                cellValue,
              );
            }
          } else if (rec.type == 0x00BD) { // MULRK
            final view = ByteData.sublistView(Uint8List.fromList(rec.data));
            int r = view.getUint16(0, Endian.little);
            int cFirst = view.getUint16(2, Endian.little);
            int numRks = (rec.data.length - 6) ~/ 6;
            for (int k = 0; k < numRks; k++) {
              int rkValue = view.getUint32(6 + k * 6, Endian.little);
              final val = _decodeRk(rkValue);
              CellValue? cellValue;
              if (val is int) {
                cellValue = IntCellValue(val);
              } else if (val is double) {
                cellValue = DoubleCellValue(val);
              }
              if (cellValue != null) {
                sheet.updateCell(
                  CellIndex.indexByColumnRow(columnIndex: cFirst + k, rowIndex: r),
                  cellValue,
                );
              }
            }
          }
        }
      }
    }

    if (!hasSheet1 && excel.tables.containsKey('Sheet1')) {
      excel.delete('Sheet1');
    }

    return excel;
  }

  List<_Record> _readRecords(List<int> stream) {
    List<_Record> records = [];
    int offset = 0;
    final view = ByteData.sublistView(Uint8List.fromList(stream));
    while (offset + 4 <= stream.length) {
      int type = view.getUint16(offset, Endian.little);
      int len = view.getUint16(offset + 2, Endian.little);
      int streamOffset = offset;
      offset += 4;
      if (offset + len > stream.length) break;
      records.add(_Record(type, stream.sublist(offset, offset + len), streamOffset));
      offset += len;
    }
    return records;
  }

  List<String> _parseSst(List<_Record> records, int sstRecordIndex) {
    final stream = _SstStream(records, sstRecordIndex);
    stream.readUint32();
    int uniqueStrings = stream.readUint32();

    List<String> strings = [];
    try {
      for (int i = 0; i < uniqueStrings; i++) {
        strings.add(_readString(stream));
      }
    } catch (_) {
      // If boundary parsing has issues, return parsed strings so far
    }
    return strings;
  }

  String _readString(_SstStream stream) {
    int charCount = stream.readUint16();
    int flags = stream.readByte();
    
    stream.currentUtf16 = (flags & 0x01) != 0;
    bool hasRichText = (flags & 0x08) != 0;
    bool hasPhonetic = (flags & 0x04) != 0;
    
    int runCount = 0;
    if (hasRichText) {
      runCount = stream.readUint16();
    }
    
    int phoneticSize = 0;
    if (hasPhonetic) {
      phoneticSize = stream.readUint32();
    }
    
    StringBuffer sb = StringBuffer();
    for (int i = 0; i < charCount; i++) {
      if (stream.currentUtf16) {
        int charCode = stream.readUint16Char();
        sb.writeCharCode(charCode);
      } else {
        int charCode = stream.readByteChar();
        sb.writeCharCode(charCode);
      }
    }
    
    for (int i = 0; i < runCount * 4; i++) {
      stream.readByte();
    }
    for (int i = 0; i < phoneticSize; i++) {
      stream.readByte();
    }
    
    return sb.toString();
  }

  dynamic _decodeRk(int rkValue) {
    int flag = rkValue & 3;
    bool isInt = (flag & 2) != 0;
    bool isX100 = (flag & 1) != 0;

    if (isInt) {
      int val = (rkValue >> 2).toSigned(30);
      return isX100 ? (val / 100.0) : val;
    } else {
      int highBits = rkValue & 0xFFFFFFFC;
      final bd = ByteData(8);
      bd.setUint32(0, 0, Endian.little);
      bd.setUint32(4, highBits, Endian.little);
      double val = bd.getFloat64(0, Endian.little);
      return isX100 ? (val / 100.0) : val;
    }
  }
}

class _BiffSheet {
  final String name;
  final int offset;
  _BiffSheet(this.name, this.offset);
}

class _Record {
  final int type;
  final List<int> data;
  final int streamOffset;
  _Record(this.type, this.data, this.streamOffset);
}

class _SstStream {
  final List<_Record> records;
  int recordIndex;
  int offset;
  bool currentUtf16 = false;
  
  _SstStream(this.records, this.recordIndex) : offset = 0;

  int readRawByte() {
    if (offset >= records[recordIndex].data.length) {
      throw StateError("Out of bounds");
    }
    return records[recordIndex].data[offset++];
  }

  bool isAtRecordEnd() {
    return offset >= records[recordIndex].data.length;
  }

  void moveToNextRecord() {
    recordIndex++;
    offset = 0;
    if (recordIndex >= records.length || records[recordIndex].type != 0x003C) {
      throw StateError("Expected CONTINUE record");
    }
  }

  int readByte() {
    if (isAtRecordEnd()) {
      moveToNextRecord();
    }
    return readRawByte();
  }

  int readByteChar() {
    if (isAtRecordEnd()) {
      moveToNextRecord();
      int continueFlags = readRawByte();
      currentUtf16 = (continueFlags & 0x01) != 0;
    }
    return readRawByte();
  }

  int readUint16Char() {
    int b1 = readByteChar();
    int b2 = readByteChar();
    return b1 | (b2 << 8);
  }

  int readUint16() {
    int b1 = readByte();
    int b2 = readByte();
    return b1 | (b2 << 8);
  }

  int readUint32() {
    int b1 = readByte();
    int b2 = readByte();
    int b3 = readByte();
    int b4 = readByte();
    return b1 | (b2 << 8) | (b3 << 16) | (b4 << 24);
  }
}
