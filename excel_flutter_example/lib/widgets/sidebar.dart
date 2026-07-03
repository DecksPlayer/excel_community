import 'package:flutter/material.dart';
import '../models/section_detail.dart';

class Sidebar extends StatelessWidget {
  final SelectedSection selectedSection;
  final ValueChanged<SelectedSection> onSectionSelected;

  const Sidebar({
    super.key,
    required this.selectedSection,
    required this.onSectionSelected,
  });

  @override
  Widget build(BuildContext context) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        const Padding(
          padding: EdgeInsets.fromLTRB(20, 20, 20, 8),
          child: Text(
            'EXCEL SPREADSHEETS',
            style: TextStyle(
              fontSize: 10,
              fontWeight: FontWeight.bold,
              color: Color(0xFF94A3B8), // Slate 400
              letterSpacing: 1.2,
            ),
          ),
        ),
        Expanded(
          child: ListView(
            padding: const EdgeInsets.symmetric(horizontal: 8),
            children: [
              _buildSidebarItem(
                context,
                SelectedSection.about,
                'About excel_community',
                Icons.info_outline,
                Colors.green,
              ),
              const Divider(color: Color(0xFFF1F5F9), height: 16),
              _buildSidebarItem(
                context,
                SelectedSection.simpleExcel,
                'Quick Start (No Chart)',
                Icons.bolt,
                Colors.amber,
              ),
              const Padding(
                padding: EdgeInsets.fromLTRB(12, 16, 12, 8),
                child: Text(
                  'Chart Types',
                  style: TextStyle(
                    fontWeight: FontWeight.bold,
                    fontSize: 11,
                    color: Color(0xFF64748B),
                  ),
                ),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.columnChart,
                'Column Chart',
                Icons.bar_chart,
                Colors.blue,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.lineChart,
                'Line Chart',
                Icons.show_chart,
                Colors.orange,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.pieChart,
                'Pie Chart',
                Icons.pie_chart,
                Colors.red,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.areaChart,
                'Area Chart',
                Icons.area_chart,
                Colors.purple,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.doughnutChart,
                'Doughnut Chart',
                Icons.donut_large,
                Colors.teal,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.radarChart,
                'Radar Chart',
                Icons.radar,
                Colors.indigo,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.barChart,
                'Bar Chart',
                Icons.horizontal_split,
                Colors.pink,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.scatterChart,
                'Scatter Chart',
                Icons.scatter_plot,
                Colors.deepOrange,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.imageEmbedding,
                'Image Embedding',
                Icons.image_outlined,
                Colors.teal,
              ),
              const Padding(
                padding: EdgeInsets.fromLTRB(12, 16, 12, 8),
                child: Text(
                  'Style & Formatting',
                  style: TextStyle(
                    fontWeight: FontWeight.bold,
                    fontSize: 11,
                    color: Color(0xFF64748B),
                  ),
                ),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.textStyles,
                'Text Underlines & Fills',
                Icons.format_underlined,
                Colors.deepPurple,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.numberFormats,
                'Number Formatting',
                Icons.pin,
                Colors.teal,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.multiSheets,
                'Multi-Worksheets',
                Icons.layers_outlined,
                Colors.cyan,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.mergedCells,
                'Merged Cells (Multi-Sheet)',
                Icons.merge_type_outlined,
                Colors.indigo,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.pivotTemplate,
                'Templates & Pivot Tables',
                Icons.content_paste_go_outlined,
                Colors.indigo,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.cellLocking,
                'Sheet Protection & Locks',
                Icons.lock_outline,
                Colors.red,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.freezePanes,
                'Freeze Panes',
                Icons.view_headline_outlined,
                Colors.indigo,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.multiFreezePanes,
                'Multi-Sheet Freeze Panes',
                Icons.layers_outlined,
                Colors.indigo,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.hiddenColumns,
                'Hidden Columns & Rows',
                Icons.visibility_off_outlined,
                Colors.teal,
              ),
              const Padding(
                padding: EdgeInsets.fromLTRB(12, 16, 12, 8),
                child: Text(
                  'Demonstrations',
                  style: TextStyle(
                    fontWeight: FontWeight.bold,
                    fontSize: 11,
                    color: Color(0xFF64748B),
                  ),
                ),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.multiPageCharts,
                'Charts on Multiple Sheets',
                Icons.auto_graph,
                Colors.deepPurple,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.allCharts,
                'All 8 Charts Grid',
                Icons.grid_view,
                Colors.blueGrey,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.fullDemo,
                'Full Sheet Report',
                Icons.star,
                Colors.amber.shade700,
              ),
            ],
          ),
        ),
      ],
    );
  }

  Widget _buildSidebarItem(
    BuildContext context,
    SelectedSection section,
    String label,
    IconData icon,
    Color color,
  ) {
    final isSelected = selectedSection == section;
    return Container(
      margin: const EdgeInsets.only(bottom: 2),
      decoration: BoxDecoration(
        color: isSelected ? color.withOpacity(0.08) : Colors.transparent,
        borderRadius: BorderRadius.circular(8),
      ),
      child: ListTile(
        onTap: () {
          onSectionSelected(section);
          if (Scaffold.of(context).isDrawerOpen) {
            Navigator.pop(context);
          }
        },
        leading: Icon(
          icon,
          color: isSelected ? color : const Color(0xFF64748B),
          size: 18,
        ),
        title: Text(
          label,
          style: TextStyle(
            color: isSelected
                ? const Color(0xFF0F172A)
                : const Color(0xFF475569),
            fontWeight: isSelected ? FontWeight.bold : FontWeight.normal,
            fontSize: 13,
          ),
        ),
        dense: true,
        visualDensity: const VisualDensity(vertical: -2),
      ),
    );
  }
}
