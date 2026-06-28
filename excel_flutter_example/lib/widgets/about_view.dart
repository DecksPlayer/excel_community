import 'package:flutter/material.dart';
import 'package:flutter_svg/flutter_svg.dart';

class AboutView extends StatelessWidget {
  const AboutView({super.key});

  @override
  Widget build(BuildContext context) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Container(
          width: double.infinity,
          padding: const EdgeInsets.symmetric(horizontal: 32, vertical: 48),
          decoration: BoxDecoration(
            gradient: const LinearGradient(
              colors: [Color(0xFF0F172A), Color(0xFF1E293B)],
              begin: Alignment.topLeft,
              end: Alignment.bottomRight,
            ),
            borderRadius: BorderRadius.circular(20),
            boxShadow: const [
              BoxShadow(
                color: Colors.black12,
                blurRadius: 10,
                offset: Offset(0, 4),
              ),
            ],
          ),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Container(
                padding: const EdgeInsets.symmetric(
                  horizontal: 12,
                  vertical: 6,
                ),
                decoration: BoxDecoration(
                  color: const Color(0xFF10B981).withOpacity(0.2),
                  borderRadius: BorderRadius.circular(20),
                  border: Border.all(
                    color: const Color(0xFF10B981),
                    width: 1,
                  ),
                ),
                child: const Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    Icon(Icons.star, color: Color(0xFF10B981), size: 14),
                    SizedBox(width: 6),
                    Text(
                      'Official Dart & Flutter Package',
                      style: TextStyle(
                        color: Color(0xFF34D399),
                        fontSize: 11,
                        fontWeight: FontWeight.bold,
                      ),
                    ),
                  ],
                ),
              ),
              const SizedBox(height: 20),
              Wrap(
                alignment: WrapAlignment.start,
                crossAxisAlignment: WrapCrossAlignment.center,
                spacing: 12,
                runSpacing: 8,
                children: [
                  SvgPicture.asset('assets/logo.svg', width: 44, height: 44),
                  const Text(
                    'excel_community',
                    style: TextStyle(
                      color: Colors.white,
                      fontSize: 36,
                      fontWeight: FontWeight.w800,
                      letterSpacing: -0.5,
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 8),
              const Text(
                'Advanced Spreadsheet Engine for Dart & Flutter',
                style: TextStyle(
                  color: Color(0xFF94A3B8),
                  fontSize: 16,
                  fontWeight: FontWeight.w500,
                ),
              ),
              const SizedBox(height: 16),
              const Text(
                'A lightweight, pure-Dart package to dynamically read and write Excel .xlsx files. '
                'It features complete ECMA-376 compliance, zero platform dependencies, robust formula parsing, cell formatting, '
                'and advanced vector-based chart anchoring. Perfect for Flutter Web (Wasm-ready), Desktop, Mobile, and Server VM.',
                style: TextStyle(
                  color: Color(0xFFCBD5E1),
                  fontSize: 13,
                  height: 1.6,
                ),
              ),
            ],
          ),
        ),
        const SizedBox(height: 32),
        const Text(
          'Key Technical Features',
          style: TextStyle(
            fontSize: 18,
            fontWeight: FontWeight.bold,
            color: Color(0xFF0F172A),
          ),
        ),
        const SizedBox(height: 16),
        GridView.count(
          crossAxisCount: MediaQuery.of(context).size.width >= 1200 ? 4 : 2,
          shrinkWrap: true,
          physics: const NeverScrollableScrollPhysics(),
          crossAxisSpacing: 16,
          mainAxisSpacing: 16,
          childAspectRatio: 1.3,
          children: [
            _buildPillarCard(
              context,
              icon: Icons.code_rounded,
              title: '100% Pure Dart',
              description:
                  'No local Excel installations, third-party bindings, or native bridges needed. Runs on all OS platforms.',
              color: Colors.blue,
            ),
            _buildPillarCard(
              context,
              icon: Icons.bar_chart_rounded,
              title: '8 Chart Types',
              description:
                  'Generates Column, Bar, Line, Area, Pie, Doughnut, Radar, and Scatter charts conforming to open spreadsheet standards.',
              color: Colors.green,
            ),
            _buildPillarCard(
              context,
              icon: Icons.style_rounded,
              title: 'Style System',
              description:
                  'Complete control over custom text sizes, bold/italic, double underlines, cell borders, and hex coloring.',
              color: Colors.purple,
            ),
            _buildPillarCard(
              context,
              icon: Icons.layers_rounded,
              title: 'Wasm & Web Ready',
              description:
                  'Optimized compilation paths that avoid file-system bottlenecks, allowing instant in-browser spreadsheet generation.',
              color: Colors.teal,
            ),
          ],
        ),
        const SizedBox(height: 32),
        Card(
          elevation: 0,
          shape: RoundedRectangleBorder(
            borderRadius: BorderRadius.circular(16),
            side: const BorderSide(color: Color(0xFFE2E8F0)),
          ),
          color: const Color(0xFF0F172A),
          child: Padding(
            padding: const EdgeInsets.all(24.0),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                const Row(
                  children: [
                    Icon(Icons.terminal, color: Color(0xFF38BDF8), size: 18),
                    SizedBox(width: 8),
                    Text(
                      'Getting Started in 1 Minute',
                      style: TextStyle(
                        color: Colors.white,
                        fontWeight: FontWeight.bold,
                        fontSize: 14,
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 12),
                const Text(
                  'Add the dependency and start writing sheets with a simple Dart script:',
                  style: TextStyle(color: Color(0xFF94A3B8), fontSize: 13),
                ),
                const SizedBox(height: 16),
                Container(
                  width: double.infinity,
                  padding: const EdgeInsets.all(16),
                  decoration: BoxDecoration(
                    color: const Color(0xFF020617),
                    borderRadius: BorderRadius.circular(8),
                  ),
                  child: const SelectableText(
                    '// 1. Create workbook & sheet\n'
                    'var excel = Excel.createExcel();\n'
                    'var sheet = excel[\'Sheet1\'];\n\n'
                    '// 2. Insert cell values & style\n'
                    'sheet.updateCell(\n'
                    '  CellIndex.indexByString("A1"), \n'
                    '  TextCellValue("Hello World"),\n'
                    '  cellStyle: CellStyle(bold: true, fontSize: 12)\n'
                    ');\n\n'
                    '// 3. Save workbook\n'
                    'var bytes = excel.save(fileName: "hello_world.xlsx");',
                    style: TextStyle(
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
        ),
      ],
    );
  }

  Widget _buildPillarCard(
    BuildContext context, {
    required IconData icon,
    required String title,
    required String description,
    required Color color,
  }) {
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
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Container(
              padding: const EdgeInsets.all(8),
              decoration: BoxDecoration(
                color: color.withOpacity(0.1),
                borderRadius: BorderRadius.circular(10),
              ),
              child: Icon(icon, color: color, size: 20),
            ),
            const SizedBox(height: 12),
            Text(
              title,
              style: const TextStyle(
                fontSize: 14,
                fontWeight: FontWeight.bold,
                color: Color(0xFF0F172A),
              ),
            ),
            const SizedBox(height: 6),
            Expanded(
              child: Text(
                description,
                style: const TextStyle(
                  fontSize: 11,
                  color: Color(0xFF64748B),
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
