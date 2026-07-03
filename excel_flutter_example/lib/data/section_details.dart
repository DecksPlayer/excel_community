import 'package:flutter/material.dart';
import '../models/section_detail.dart';
import 'code_snippets.dart';

SectionDetail getSectionDetail(SelectedSection section) {
  switch (section) {
    case SelectedSection.about:
      return SectionDetail(
        title: 'About the Project',
        description:
            'A pure-Dart engine to create and read Excel spreadsheets dynamically.',
        icon: Icons.info_outline,
        themeColor: Colors.green,
        highlights: [],
        codeSnippet: '',
      );
    case SelectedSection.simpleExcel:
      return SectionDetail(
        title: 'Quick Start Excel',
        description:
            'A basic Excel spreadsheet containing formatted texts and integer values without any charts.',
        icon: Icons.bolt,
        themeColor: Colors.amber,
        highlights: [
          'Basic single-sheet workbook creation',
          'Supports Text and Integer cell values',
          'Fast saving using File (Desktop) or auto-downloads (Web)',
        ],
        codeSnippet: simpleExcelSnippet,
      );
    case SelectedSection.columnChart:
      return SectionDetail(
        title: 'Column Chart',
        description: 'Vertical clustered columns for category data comparison.',
        icon: Icons.bar_chart,
        themeColor: Colors.blue,
        highlights: [
          'Professional standard Excel Column charts',
          'Compliance with ECMA-376 OOXML standards',
          'Supports multiple data series overlay',
        ],
        codeSnippet: columnChartSnippet,
      );
    case SelectedSection.lineChart:
      return SectionDetail(
        title: 'Line Chart',
        description:
            'A trend display over continuous intervals with circular series markers.',
        icon: Icons.show_chart,
        themeColor: Colors.orange,
        highlights: [
          'Thick line borders (28575 EMUs) for modern aesthetics',
          'Circular data markers matching series line colors',
          'Perfect for time-series and continuous trends',
        ],
        codeSnippet: lineChartSnippet,
      );
    case SelectedSection.pieChart:
      return SectionDetail(
        title: 'Pie Chart',
        description:
            'Circular slice distribution demonstrating relative percentages.',
        icon: Icons.pie_chart,
        themeColor: Colors.red,
        highlights: [
          'Supports single series segment visualization',
          'Shuffles standard 20-color palette for unique slices',
          'Ensures no repeated colors within the same pie chart',
        ],
        codeSnippet: pieChartSnippet,
      );
    case SelectedSection.areaChart:
      return SectionDetail(
        title: 'Area Chart',
        description:
            'Filled area series overlay showing accumulation over categories.',
        icon: Icons.area_chart,
        themeColor: Colors.purple,
        highlights: [
          'Semi-transparent fills (50% opacity) for overlapping series visibility',
          'Opaque contours (90% opacity, 2.25pt thick) defining series paths',
          'Clean vector output for modern reporting sheets',
        ],
        codeSnippet: areaChartSnippet,
      );
    case SelectedSection.doughnutChart:
      return SectionDetail(
        title: 'Doughnut Chart',
        description:
            'A circular chart with a central cutout, displaying part-to-whole data.',
        icon: Icons.donut_large,
        themeColor: Colors.teal,
        highlights: [
          'Built-in hole size set to 50% for optimal readability',
          'Shuffled colors from a professional 20-color palette',
          'Compatible with Microsoft Excel, Excel 365, and Google Sheets',
        ],
        codeSnippet: doughnutChartSnippet,
      );
    case SelectedSection.radarChart:
      return SectionDetail(
        title: 'Radar Chart',
        description:
            'Spider chart showing multi-variable comparison in a radial layout.',
        icon: Icons.radar,
        themeColor: Colors.indigo,
        highlights: [
          'Filled radar series support with 45% transparency layers',
          'Bold contour boundary lines with 85% opacity settings',
          'Clean radial axes grid generation compliant with OOXML standards',
        ],
        codeSnippet: radarChartSnippet,
      );
    case SelectedSection.barChart:
      return SectionDetail(
        title: 'Bar Chart',
        description:
            'Horizontal clustered bars ideal for displaying descriptive labels.',
        icon: Icons.horizontal_split,
        themeColor: Colors.pink,
        highlights: [
          'Horizontal clustered columns formatting',
          'Consistent 12-color rotating series palette',
          'Thin borders matching series color fills (0.75pt)',
        ],
        codeSnippet: barChartSnippet,
      );
    case SelectedSection.scatterChart:
      return SectionDetail(
        title: 'Scatter Chart',
        description:
            'XY plotting representing numerical values mapping correlations.',
        icon: Icons.scatter_plot,
        themeColor: Colors.deepOrange,
        highlights: [
          'Requires both numeric X values and numeric Y values',
          'Solid data points highlighted by a clean white outer halo',
          'Aids readability when dealing with heavy overlapping point maps',
        ],
        codeSnippet: scatterChartSnippet,
      );
    case SelectedSection.imageEmbedding:
      return SectionDetail(
        title: 'Image Embedding',
        description:
            'Embed PNG, JPEG, SVG, WebP, and other image formats into your worksheets.',
        icon: Icons.image_outlined,
        themeColor: Colors.green,
        highlights: [
          'Embed images with custom pixel dimensions',
          'Supports positioning using column/row coordinates and EMU offsets',
          'Compatible with PNG, JPEG, BMP, GIF, SVG, and WebP formats',
        ],
        codeSnippet: imageEmbeddingSnippet,
      );
    case SelectedSection.textStyles:
      return SectionDetail(
        title: 'Text Styles & Underlines',
        description:
            'Format fonts, sizes, italic/bold, colors, backgrounds, strikethrough, and underlines.',
        icon: Icons.format_underlined,
        themeColor: Colors.deepPurple,
        highlights: [
          'Underline options: Underline.None, Underline.Single, Underline.Double',
          'Font family selection (Arial, Calibri, Comic Sans, etc.)',
          'Text strikethrough styling and color preservation',
        ],
        codeSnippet: textStylesSnippet,
      );
    case SelectedSection.numberFormats:
      return SectionDetail(
        title: 'Number Formatting',
        description: 'Built-in standard and custom number formatting rules.',
        icon: Icons.pin,
        themeColor: Colors.teal,
        highlights: [
          'Integrates built-in ECMA-376 Standard formats (0-44)',
          'Formats float decimals, scientific exponents, percentages, currencies',
          'Accounting currency values with indent alignment support (ID 44)',
        ],
        codeSnippet: numberFormatsSnippet,
      );
    case SelectedSection.multiSheets:
      return SectionDetail(
        title: 'Multi-Worksheets',
        description:
            'Create and organize multiple sheets inside a single Excel workbook.',
        icon: Icons.layers_outlined,
        themeColor: Colors.cyan,
        highlights: [
          'Create, rename, and select sheets on the fly',
          'Delete default template sheets to clean up output',
          'Isolate data schemas across dedicated worksheets',
        ],
        codeSnippet: multiSheetsSnippet,
      );
    case SelectedSection.mergedCells:
      return SectionDetail(
        title: 'Merged Cells (Multi-Sheet)',
        description:
            'Merge rows and columns dynamically across multiple worksheets while applying custom cell styles, alignment, and formatting.',
        icon: Icons.merge_type_outlined,
        themeColor: Colors.indigo,
        highlights: [
          'Merge multiple columns or rows using sheet.merge(start, end)',
          'Customize text alignment (Horizontal/Vertical) inside merged ranges',
          'Add custom borders, font styles, and background fills to merged headers',
          'Ensure values are stored correctly at the top-left origin cell',
        ],
        codeSnippet: mergedCellsSnippet,
      );
    case SelectedSection.pivotTemplate:
      return SectionDetail(
        title: 'Templates & Pivot Tables',
        description:
            'Modify pre-configured Excel templates containing Pivot Tables while preserving all formulas and connections.',
        icon: Icons.content_paste_go_outlined,
        themeColor: Colors.indigo,
        highlights: [
          'Load templates using Excel.decodeBytes(...)',
          'All Pivot Cache Definitions and Filters are fully preserved',
          'Modify source data sheets, and let Excel calculate Pivot updates on open',
        ],
        codeSnippet: pivotTemplateSnippet,
      );
    case SelectedSection.cellLocking:
      return SectionDetail(
        title: 'Sheet Protection & Locks',
        description:
            'Configure sheet protection and custom cell locks programmatically. Mark specific input ranges as editable while locking down headers or keys.',
        icon: Icons.lock_outline,
        themeColor: Colors.red,
        highlights: [
          'Enable worksheet protection programmatically using sheet.protect(password)',
          'Customize CellStyle(locked: true) for read-only rows/headers',
          'Configure CellStyle(locked: false) to keep input cells fully editable',
        ],
        codeSnippet: cellLockingSnippet,
      );
    case SelectedSection.freezePanes:
      return SectionDetail(
        title: 'Freeze Panes',
        description:
            'Freeze rows and columns programmatically to keep headers or keys always visible while scrolling.',
        icon: Icons.view_headline_outlined,
        themeColor: Colors.indigo,
        highlights: [
          'Support for custom frozen rows and columns properties',
          'Row 1 and Column A frozen programmatically using sheet.frozenRows/Columns',
          'Keeps headers/first columns fixed in place during vertical & horizontal scrolling',
        ],
        codeSnippet: freezePanesSnippet,
      );
    case SelectedSection.multiFreezePanes:
      return SectionDetail(
        title: 'Multi-Worksheet Freeze Panes',
        description:
            'Apply different freeze pane configurations per sheet in the same workbook. Demonstrates how Sales, Inventory, Customers, and Logs each get their own <pane> element.',
        icon: Icons.layers_outlined,
        themeColor: Colors.indigo,
        highlights: [
          'Four sheets in a single workbook, each with its own frozenRows / frozenColumns',
          'Combines all four modes: rows+cols, rows only, columns only, and no freeze',
          'Round-trip verified: decoding the file restores every freeze pane',
        ],
        codeSnippet: multiFreezePanesSnippet,
      );
    case SelectedSection.hiddenColumns:
      return SectionDetail(
        title: 'Hidden Columns & Rows',
        description:
            'Hide specific columns and rows in a worksheet to mask confidential details or keep internal formula inputs invisible. Column C (Base Salary) and Row 6 (Inactive employee) are hidden.',
        icon: Icons.visibility_off,
        themeColor: Colors.teal,
        highlights: [
          'Columns can be set to hidden using sheet.setColumnHidden(index, true)',
          'Rows can be set to hidden using sheet.setRowHidden(index, true)',
          'Values are preserved, but hidden from initial display in Excel',
        ],
        codeSnippet: hiddenColumnsSnippet,
      );
    case SelectedSection.multiPageCharts:
      return SectionDetail(
        title: 'Charts on Multiple Sheets',
        description:
            'One chart per worksheet — Sales (Column), Market Share (Pie), Trends (Line), Performance (Radar), Correlation (Scatter). Validates that charts render correctly on non-default sheets.',
        icon: Icons.auto_graph,
        themeColor: Colors.deepPurple,
        highlights: [
          '5 sheets, each with a distinct chart type',
          'Each sheet owns an independent drawing + rels file',
          'Charts on non-default sheets display correctly in Excel',
        ],
        codeSnippet: multiPageChartsSnippet,
      );
    case SelectedSection.allCharts:
      return SectionDetail(
        title: 'All Charts Grid',
        description:
            'A composite grid layout showing all 8 supported charts side-by-side.',
        icon: Icons.grid_view,
        themeColor: Colors.blueGrey,
        highlights: [
          'Grid sheet layout positioning charts dynamically',
          'Leverages distinct custom styles for each chart builder type',
          'Full compliance validation under a single workbook relationship XML',
        ],
        codeSnippet: allChartsSnippet,
      );
    case SelectedSection.fullDemo:
      return SectionDetail(
        title: 'Full Sheet Demo',
        description:
            'A complete real-world reporting sheet with headers, custom cell styles, formulas, and charts.',
        icon: Icons.star,
        themeColor: Colors.amber.shade700,
        highlights: [
          'Alternating row styles with customized thin grid borders',
          'Emits Excel formulas: SUM(A1:A5) and arithmetic equations',
          'Anchors column charts automatically inside a styled spreadsheet',
        ],
        codeSnippet: fullDemoSnippet,
      );
  }
}
