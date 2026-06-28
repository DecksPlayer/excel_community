import 'package:flutter/material.dart';
import 'package:flutter/painting.dart' as fp;
import '../models/section_detail.dart';

class SpreadsheetPreview extends StatelessWidget {
  final SelectedSection selectedSection;
  final SectionDetail detail;

  const SpreadsheetPreview({
    super.key,
    required this.selectedSection,
    required this.detail,
  });

  @override
  Widget build(BuildContext context) {
    const double colWidth = 52.0;
    const double rowHeight = 28.0;

    return Center(
      child: Container(
        width: 32 + colWidth * 6,
        decoration: BoxDecoration(
          border: fp.Border.all(color: const Color(0xFFCBD5E1)),
          borderRadius: BorderRadius.circular(8),
          color: Colors.white,
        ),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          mainAxisSize: MainAxisSize.min,
          children: [
            // Header Row (A, B, C, D...)
            Container(
              color: const Color(0xFFF1F5F9),
              height: 24,
              child: Row(
                children: [
                  Container(
                    width: 32,
                    decoration: const BoxDecoration(
                      border: fp.Border(
                        right: fp.BorderSide(color: Color(0xFFCBD5E1)),
                        bottom: fp.BorderSide(color: Color(0xFFCBD5E1)),
                      ),
                    ),
                  ),
                  ...List.generate(6, (i) {
                    final colName = String.fromCharCode(65 + i);
                    return Container(
                      width: colWidth,
                      alignment: Alignment.center,
                      decoration: const BoxDecoration(
                        border: fp.Border(
                          right: fp.BorderSide(color: Color(0xFFCBD5E1)),
                          bottom: fp.BorderSide(color: Color(0xFFCBD5E1)),
                        ),
                      ),
                      child: Text(
                        colName,
                        style: const TextStyle(
                          fontSize: 10,
                          fontWeight: FontWeight.bold,
                          color: Color(0xFF64748B),
                        ),
                      ),
                    );
                  }),
                ],
              ),
            ),
            // Data Rows (1 to 10)
            SizedBox(
              height: rowHeight * 10,
              child: Stack(
                children: [
                  Column(
                    children: List.generate(10, (rowIndex) {
                      final rowNum = rowIndex + 1;
                      return SizedBox(
                        height: rowHeight,
                        child: Row(
                          children: [
                            Container(
                              width: 32,
                              color: const Color(0xFFF1F5F9),
                              alignment: Alignment.center,
                              decoration: const BoxDecoration(
                                border: fp.Border(
                                  right: fp.BorderSide(color: Color(0xFFCBD5E1)),
                                  bottom: fp.BorderSide(color: Color(0xFFCBD5E1)),
                                ),
                              ),
                              child: Text(
                                '$rowNum',
                                style: const TextStyle(
                                  fontSize: 10,
                                  fontWeight: FontWeight.bold,
                                  color: Color(0xFF64748B),
                                ),
                              ),
                            ),
                            ...List.generate(6, (colIndex) {
                              return Container(
                                width: colWidth,
                                decoration: const BoxDecoration(
                                  border: fp.Border(
                                    right: fp.BorderSide(color: Color(0xFFE2E8F0)),
                                    bottom: fp.BorderSide(color: Color(0xFFE2E8F0)),
                                  ),
                                ),
                                child: _buildMockCellContent(rowIndex, colIndex),
                              );
                            }),
                          ],
                        ),
                      );
                    }),
                  ),
                  _buildMockOverlay(context, colWidth, rowHeight),
                ],
              ),
            ),
            if (selectedSection == SelectedSection.multiSheets) ...[
              const Divider(height: 1, color: Color(0xFFCBD5E1)),
              Container(
                color: const Color(0xFFF8FAFC),
                height: 28,
                padding: const EdgeInsets.symmetric(horizontal: 8),
                child: Row(
                  children: [
                    const Icon(Icons.playlist_add, size: 14, color: Color(0xFF64748B)),
                    const SizedBox(width: 8),
                    _buildSheetTab("Summary", isActive: true),
                    _buildSheetTab("Revenues", isActive: false),
                    _buildSheetTab("Expenses", isActive: false),
                  ],
                ),
              ),
            ],
          ],
        ),
      ),
    );
  }

  Widget _buildSheetTab(String name, {required bool isActive}) {
    return Container(
      margin: const EdgeInsets.only(top: 4, right: 4),
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
      decoration: BoxDecoration(
        color: isActive ? Colors.white : Colors.transparent,
        borderRadius: const BorderRadius.only(
          topLeft: Radius.circular(4),
          topRight: Radius.circular(4),
        ),
        border: isActive
            ? const fp.Border(
                left: fp.BorderSide(color: Color(0xFFCBD5E1)),
                right: fp.BorderSide(color: Color(0xFFCBD5E1)),
                top: fp.BorderSide(color: Color(0xFF10B981), width: 2),
              )
            : null,
      ),
      child: Text(
        name,
        style: TextStyle(
          fontSize: 9,
          fontWeight: isActive ? FontWeight.bold : FontWeight.normal,
          color: isActive ? const Color(0xFF0F172A) : const Color(0xFF64748B),
        ),
      ),
    );
  }

  Widget _buildMockCellContent(int rowIndex, int colIndex) {
    if (selectedSection == SelectedSection.multiSheets) {
      if (rowIndex == 0 && colIndex == 0) {
        return Container(
          color: Colors.blue.shade50,
          alignment: Alignment.center,
          child: const Text('Title', style: TextStyle(fontSize: 8, fontWeight: FontWeight.bold)),
        );
      }
      if (rowIndex == 2 && colIndex == 0) {
        return const Padding(
          padding: EdgeInsets.only(left: 4, top: 4),
          child: Text('Summary Sheet data is placed here.', style: TextStyle(fontSize: 6)),
        );
      }
      if (rowIndex == 4 && colIndex == 0) {
        return const Padding(
          padding: EdgeInsets.only(left: 4),
          child: Text('Check tabs at the bottom!', style: TextStyle(fontSize: 6, fontWeight: FontWeight.bold, color: Colors.blue)),
        );
      }
    } else if (selectedSection == SelectedSection.imageEmbedding) {
      if (rowIndex == 0 && colIndex == 0) {
        return Container(
          color: Colors.teal.shade50,
          alignment: Alignment.center,
          child: const Text('Title', style: TextStyle(fontSize: 8, fontWeight: FontWeight.bold)),
        );
      }
      if (rowIndex == 2 && colIndex == 0) {
        return const Padding(
          padding: EdgeInsets.only(left: 4, top: 4),
          child: Text('Image placed at B5:', style: TextStyle(fontSize: 6)),
        );
      }
    } else if (selectedSection == SelectedSection.textStyles) {
      if (rowIndex == 0 && colIndex == 0) {
        return Container(
          color: Colors.blue.shade50,
          alignment: Alignment.center,
          child: const Text('Title', style: TextStyle(fontSize: 8, fontWeight: FontWeight.bold)),
        );
      }
      if (rowIndex == 2 && colIndex == 1) {
        return Container(
          alignment: Alignment.center,
          child: const Text('Normal', style: TextStyle(fontSize: 7)),
        );
      }
      if (rowIndex == 3 && colIndex == 1) {
        return Container(
          alignment: Alignment.center,
          child: const Text('Underline', style: TextStyle(fontSize: 7, decoration: TextDecoration.underline)),
        );
      }
      if (rowIndex == 4 && colIndex == 1) {
        return Container(
          alignment: Alignment.center,
          child: const Text('Double', style: TextStyle(fontSize: 7, decoration: TextDecoration.underline, decorationStyle: TextDecorationStyle.double)),
        );
      }
      if (rowIndex == 6 && colIndex == 0) {
        return Container(
          alignment: Alignment.center,
          child: const Text('Strikethrough', style: TextStyle(fontSize: 7, decoration: TextDecoration.lineThrough, color: Colors.red)),
        );
      }
    } else if (selectedSection == SelectedSection.numberFormats) {
      if (rowIndex == 0) {
        return Container(
          color: Colors.teal.shade50,
          alignment: Alignment.center,
          child: Text(colIndex == 0 ? 'ID' : colIndex == 1 ? 'Desc' : colIndex == 2 ? 'Raw' : colIndex == 3 ? 'Format' : '', style: const TextStyle(fontSize: 8, fontWeight: FontWeight.bold)),
        );
      }
      if (rowIndex > 0 && rowIndex < 9) {
        if (colIndex == 0) return Center(child: Text('${rowIndex * 2}', style: const TextStyle(fontSize: 7)));
        if (colIndex == 1) return const Padding(padding: EdgeInsets.only(left: 2), child: Text('Demo fmt', style: TextStyle(fontSize: 6)));
        if (colIndex == 2) return const Center(child: Text('1234.5', style: TextStyle(fontSize: 7)));
        if (colIndex == 3) return Container(color: Colors.green.shade50, child: const Center(child: Text('\$1,234.50', style: TextStyle(fontSize: 7, fontWeight: FontWeight.bold))));
      }
    } else {
      if (rowIndex == 0) {
        if (colIndex == 0) return Container(color: Colors.blue.shade50, child: const Center(child: Text('Label', style: TextStyle(fontSize: 8, fontWeight: FontWeight.bold))));
        if (colIndex == 1) return Container(color: Colors.blue.shade50, child: const Center(child: Text('Val 1', style: TextStyle(fontSize: 8, fontWeight: FontWeight.bold))));
        if (colIndex == 2) return Container(color: Colors.blue.shade50, child: const Center(child: Text('Val 2', style: TextStyle(fontSize: 8, fontWeight: FontWeight.bold))));
      } else if (rowIndex > 0 && rowIndex < 7) {
        if (colIndex == 0) return Center(child: Text(rowIndex == 1 ? 'Jan' : rowIndex == 2 ? 'Feb' : rowIndex == 3 ? 'Mar' : rowIndex == 4 ? 'Apr' : rowIndex == 5 ? 'May' : 'Jun', style: const TextStyle(fontSize: 7)));
        if (colIndex == 1) return Center(child: Text('${rowIndex * 10}', style: const TextStyle(fontSize: 7)));
        if (colIndex == 2) return Center(child: Text('${rowIndex * 15}', style: const TextStyle(fontSize: 7)));
      } else if (rowIndex == 7 && selectedSection == SelectedSection.fullDemo) {
        if (colIndex == 0) return Container(color: Colors.green.shade50, child: const Center(child: Text('TOTAL', style: TextStyle(fontSize: 8, fontWeight: FontWeight.bold))));
        if (colIndex == 1) return Container(color: Colors.green.shade50, child: const Center(child: Text('SUM(B2:B7)', style: TextStyle(fontSize: 6))));
        if (colIndex == 2) return Container(color: Colors.green.shade50, child: const Center(child: Text('SUM(C2:C7)', style: TextStyle(fontSize: 6))));
      }
    }
    return const SizedBox.shrink();
  }

  Widget _buildMockOverlay(BuildContext context, double colWidth, double rowHeight) {
    if (selectedSection == SelectedSection.simpleExcel ||
        selectedSection == SelectedSection.textStyles ||
        selectedSection == SelectedSection.numberFormats) {
      return const SizedBox.shrink();
    }

    if (selectedSection == SelectedSection.imageEmbedding) {
      return Positioned(
        left: 32 + colWidth * 1 + 3,
        top: rowHeight * 4 + 3,
        child: Container(
          width: colWidth * 3 - 6,
          height: rowHeight * 4 - 6,
          decoration: BoxDecoration(
            color: Colors.green.shade50,
            border: fp.Border.all(color: Colors.green, width: 1.5),
            borderRadius: BorderRadius.circular(6),
            boxShadow: const [
              BoxShadow(
                color: Colors.black12,
                blurRadius: 3,
                offset: Offset(1, 1),
              ),
            ],
          ),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              const Icon(Icons.image, size: 24, color: Colors.green),
              const SizedBox(height: 2),
              Text(
                'Embedded Image\n(Green Square)',
                textAlign: TextAlign.center,
                style: TextStyle(fontSize: 7, fontWeight: FontWeight.bold, color: Colors.green.shade900, height: 1.0),
              ),
            ],
          ),
        ),
      );
    }

    return Positioned(
      left: 32 + colWidth * 3 + 3,
      top: rowHeight * 1 + 3,
      child: Container(
        width: colWidth * 3 - 6,
        height: rowHeight * 7 - 6,
        decoration: BoxDecoration(
          color: Colors.white,
          border: fp.Border.all(color: detail.themeColor, width: 1.5),
          borderRadius: BorderRadius.circular(6),
          boxShadow: const [
            BoxShadow(
              color: Colors.black12,
              blurRadius: 3,
              offset: Offset(1, 1),
            ),
          ],
        ),
        child: Column(
          children: [
            Container(
              padding: const EdgeInsets.symmetric(vertical: 2, horizontal: 4),
              color: detail.themeColor.withOpacity(0.08),
              width: double.infinity,
              child: Row(
                children: [
                  Icon(detail.icon, size: 8, color: detail.themeColor),
                  const SizedBox(width: 3),
                  const Text(
                    'Chart Anchor',
                    style: TextStyle(fontSize: 6, fontWeight: FontWeight.bold, color: Color(0xFF1E293B)),
                  ),
                ],
              ),
            ),
            Expanded(
              child: Center(
                child: Column(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    Icon(detail.icon, size: 24, color: detail.themeColor),
                    const SizedBox(height: 2),
                    Text(
                      detail.title.replaceAll(' ', '\n'),
                      textAlign: TextAlign.center,
                      style: TextStyle(fontSize: 7, fontWeight: FontWeight.bold, color: detail.themeColor, height: 1.0),
                    ),
                  ],
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
