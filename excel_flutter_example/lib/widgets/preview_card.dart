import 'package:flutter/material.dart';
import '../models/section_detail.dart';
import 'spreadsheet_preview.dart';

class PreviewCard extends StatelessWidget {
  final SectionDetail detail;
  final String status;
  final SelectedSection selectedSection;

  const PreviewCard({
    super.key,
    required this.detail,
    required this.status,
    required this.selectedSection,
  });

  @override
  Widget build(BuildContext context) {
    return Card(
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
        side: const BorderSide(color: Color(0xFFE2E8F0)),
      ),
      color: Colors.white,
      child: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Icon(
                  Icons.remove_red_eye_outlined,
                  color: detail.themeColor,
                  size: 18,
                ),
                const SizedBox(width: 8),
                const Text(
                  'Sheet Preview & Details',
                  style: TextStyle(
                    fontWeight: FontWeight.bold,
                    fontSize: 13,
                    color: Color(0xFF0F172A),
                  ),
                ),
              ],
            ),
            const SizedBox(height: 10),
            const Divider(color: Color(0xFFF1F5F9)),
            const SizedBox(height: 6),
            const Text(
              'Spreadsheet highlights:',
              style: TextStyle(
                fontWeight: FontWeight.bold,
                fontSize: 12,
                color: Color(0xFF475569),
              ),
            ),
            const SizedBox(height: 6),
            Column(
              children: detail.highlights.map((h) {
                return Padding(
                  padding: const EdgeInsets.only(bottom: 6),
                  child: Row(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      const Icon(
                        Icons.check_circle_outline,
                        size: 14,
                        color: Color(0xFF10B981),
                      ),
                      const SizedBox(width: 6),
                      Expanded(
                        child: Text(
                          h,
                          style: const TextStyle(
                            fontSize: 11,
                            color: Color(0xFF64748B),
                          ),
                        ),
                      ),
                    ],
                  ),
                );
              }).toList(),
            ),
            const SizedBox(height: 12),
            const Text(
              'Generation Status:',
              style: TextStyle(
                fontWeight: FontWeight.bold,
                fontSize: 12,
                color: Color(0xFF475569),
              ),
            ),
            const SizedBox(height: 4),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
              decoration: BoxDecoration(
                color: const Color(0xFFF8FAFC),
                borderRadius: BorderRadius.circular(8),
                border: Border.all(color: const Color(0xFFE2E8F0)),
              ),
              child: Text(
                status,
                style: const TextStyle(
                  fontSize: 11,
                  color: Color(0xFF334155),
                  height: 1.3,
                ),
              ),
            ),
            const SizedBox(height: 16),
            const Text(
              'Mock Spreadsheet Anchoring View:',
              style: TextStyle(
                fontWeight: FontWeight.bold,
                fontSize: 12,
                color: Color(0xFF475569),
              ),
            ),
            const SizedBox(height: 10),
            const Divider(color: Color(0xFFF1F5F9)),
            const SizedBox(height: 12),
            SpreadsheetPreview(
              selectedSection: selectedSection,
              detail: detail,
            ),
          ],
        ),
      ),
    );
  }
}
