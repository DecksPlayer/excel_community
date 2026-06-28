import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import '../models/section_detail.dart';

class CodeViewCard extends StatelessWidget {
  final SectionDetail detail;

  const CodeViewCard({super.key, required this.detail});

  @override
  Widget build(BuildContext context) {
    return Card(
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
        side: const BorderSide(color: Color(0xFFE2E8F0)),
      ),
      color: const Color(0xFF0F172A), // Slate 900
      child: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                const Row(
                  children: [
                    Icon(
                      Icons.code,
                      color: Color(0xFF38BDF8),
                      size: 18,
                    ), // Sky 400
                    SizedBox(width: 8),
                    Text(
                      'Dart Source Code',
                      style: TextStyle(
                        color: Color(0xFFE2E8F0),
                        fontWeight: FontWeight.bold,
                        fontSize: 13,
                      ),
                    ),
                  ],
                ),
                TextButton.icon(
                  onPressed: () {
                    Clipboard.setData(ClipboardData(text: detail.codeSnippet));
                    ScaffoldMessenger.of(context).showSnackBar(
                      const SnackBar(
                        content: Row(
                          children: [
                            Icon(Icons.check_circle, color: Color(0xFF10B981)),
                            SizedBox(width: 8),
                            Text('Code copied to clipboard successfully!'),
                          ],
                        ),
                        backgroundColor: Color(0xFF1E293B),
                        behavior: SnackBarBehavior.floating,
                      ),
                    );
                  },
                  icon: const Icon(
                    Icons.copy,
                    size: 13,
                    color: Color(0xFF38BDF8),
                  ),
                  label: const Text(
                    'Copy Code',
                    style: TextStyle(
                      color: Color(0xFF38BDF8),
                      fontSize: 12,
                      fontWeight: FontWeight.bold,
                    ),
                  ),
                ),
              ],
            ),
            const Divider(color: Colors.white10),
            const SizedBox(height: 8),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(12),
              decoration: BoxDecoration(
                color: const Color(0xFF020617), // Slate 950
                borderRadius: BorderRadius.circular(8),
              ),
              child: SelectableText(
                detail.codeSnippet,
                style: const TextStyle(
                  fontFamily: 'monospace',
                  fontSize: 11,
                  color: Color(0xFFF1F5F9),
                  height: 1.4,
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
