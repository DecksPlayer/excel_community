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
                SelectedSection.chartDataLabels,
                'Chart Data Labels',
                Icons.label_outline,
                Colors.lightBlue,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.chartColors,
                'Chart Color Customization',
                Icons.palette,
                Colors.deepPurple,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.newCharts,
                'Bubble, Stock & Stacked',
                Icons.bubble_chart,
                Colors.teal,
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
                SelectedSection.fontsStyles,
                'Fonts & Styles',
                Icons.font_download_outlined,
                Colors.indigo,
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
                SelectedSection.cellComments,
                'Cell Comments',
                Icons.comment_outlined,
                Colors.teal,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.conditionalFormatting,
                'Conditional Formatting',
                Icons.palette_outlined,
                Colors.deepOrange,
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
                SelectedSection.readAsset,
                'Read Asset (Borders & Data)',
                Icons.folder_open_outlined,
                const Color(0xFF0284C7),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.formulasDisplayText,
                'Formulas & Display Text',
                Icons.functions,
                Colors.blueGrey,
              ),
              _buildSidebarItem(
                context,
                SelectedSection.autoFilter,
                'AutoFilter (<autoFilter>)',
                Icons.filter_alt_outlined,
                const Color(0xFF2563EB),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.tabColor,
                'Sheet Tab Colors (<tabColor>)',
                Icons.color_lens_outlined,
                const Color(0xFF10B981),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.pageSetup,
                'Page Setup & Printing (<pageSetup>)',
                Icons.print_outlined,
                const Color(0xFF0EA5E9),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.dataExport,
                'Data Export (JSON / CSV)',
                Icons.data_object,
                const Color(0xFF7C3AED),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.hyperlinks,
                'Cell Hyperlinks (<hyperlinks>)',
                Icons.link,
                const Color(0xFF2563EB),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.dataValidation,
                'Data Validation & Dropdowns',
                Icons.fact_check_outlined,
                const Color(0xFF0D9488),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.grouping,
                'Row & Column Grouping',
                Icons.account_tree_outlined,
                const Color(0xFFEA580C),
              ),
              _buildSidebarItem(
                context,
                SelectedSection.tables,
                'Excel Tables (<tableParts>)',
                Icons.table_chart_outlined,
                const Color(0xFF2563EB),
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
    // The ListTile paints its highlight and ink splashes on its own Material,
    // so they are not hidden by the decorated containers around the sidebar.
    return Padding(
      padding: const EdgeInsets.only(bottom: 2),
      child: Material(
        color: Colors.transparent,
        child: ListTile(
          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
          tileColor: isSelected ? color.withValues(alpha: 0.08) : null,
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
      ),
    );
  }
}
