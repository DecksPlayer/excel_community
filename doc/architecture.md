# Clean Code Architecture - Chart System

This document outlines the architectural principles, component layout, design patterns, and code structure of the chart generation system within the `excel_community` library.

---

## 📋 Principles Applied

✅ **SOLID Principles**
- **S**ingle Responsibility: Each class has one reason to change.
- **O**pen/Closed: Open for extension, closed for modification.
- **L**iskov Substitution: Builders are interchangeable.
- **I**nterface Segregation: Simple, focused interfaces.
- **D**ependency Inversion: High-level modules depend on abstractions, not concrete implementations.

---

## 🏗️ Architecture Overview

```
excel_community/
├── lib/src/utilities/
│   ├── chart_xml_writer.dart          # Orchestrator (high-level logic)
│   ├── chart_color_config.dart        # Color configuration (centralized)
│   └── chart_builders/                # Builder pattern implementation
│       ├── chart_style_builder.dart            # Base interface
│       ├── column_bar_chart_builder.dart       # Column/Bar charts
│       ├── line_chart_builder.dart             # Line charts
│       ├── area_chart_builder.dart             # Area charts (transparency)
│       ├── scatter_chart_builder.dart          # Scatter charts
│       ├── pie_chart_builder.dart              # Pie/Doughnut charts
│       ├── radar_chart_builder.dart            # Radar charts
│       └── chart_style_builder_factory.dart    # Factory pattern
```

---

## 📦 Component Responsibilities

### 1. `ChartXmlWriter` (Orchestrator)
- **Responsibility**: Coordinates high-level XML structure generation.
- **Size**: ~200 lines.
- **Does**: 
  - XML structure mapping (Drawing & Chart files).
  - Delegates styling definitions to specific chart builders.
- **Doesn't**: 
  - Manage colors directly.
  - Implement chart-specific rendering rules.

### 2. `ChartColorConfig` (Configuration)
- **Responsibility**: Centralizes color management and palettes.
- **Size**: ~100 lines.
- **Does**:
  - Stores standard series palettes, radar charts, and randomized pie charts.
  - Provides color accessors.
  - Controls styling parameters (such as line widths and opacities).
- **Doesn't**:
  - Direct XML building.
  - Handle chart layouts.

### 3. `ChartStyleBuilder` (Interface)
- **Responsibility**: Defines the builder contract.
- **Size**: ~15 lines.
- **Methods**:
  - `buildProperties()`: Inject chart-specific XML properties.
  - `buildSeriesStyle()`: Apply series-specific styles and colors.

### 4. Individual Builders (7 classes)
Each builder class manages exactly ONE chart type:
- **ColumnBarChartBuilder**: Column and Bar chart styling (solid fills, matching borders, clustering).
- **LineChartBuilder**: Line charts (thick contours, circular markers).
- **AreaChartBuilder**: Area charts (90% line opacity, 50% fill opacity).
- **ScatterChartBuilder**: Scatter plots (solid fill markers with a distinct white border).
- **PieChartBuilder**: Pie & Doughnut charts (shuffled colors per slice, hole sizing, slice offset angle).
- **RadarChartBuilder**: Radar charts (opacity matching, line-only vs. filled variants).

### 5. `ChartStyleBuilderFactory` (Factory)
- **Responsibility**: Instantiates the correct builder based on chart type.
- **Pattern**: Factory Method.

---

## 🎯 Benefits of This Architecture

### ✅ Single Responsibility
Each class has one reason to change:
- `ChartColorConfig`: Adjust color schemes.
- `PieChartBuilder`: Modify pie rendering features.
- `RadarChartBuilder`: Customize radar grids.

### ✅ Easy to Extend
Adding a new chart type (e.g. Waterfall Chart) is simple:
1. Create a new builder class `waterfall_chart_builder.dart` implementing the `ChartStyleBuilder` interface.
2. Register it in `ChartStyleBuilderFactory`.
3. Safe execution: no modifications are required in other builders.

### ✅ Easy to Test
Each builder can be isolated and verified independently:
```dart
test('PieChartBuilder generates random colors', () {
  final builder = PieChartBuilder();
  // Test only pie chart logic
});
```

---

## 📊 Comparison: Before vs. After

### ❌ Before (Monolithic)
```
chart_xml_writer.dart: 550 lines
├── Everything mixed together
├── Hard to find specific logic
├── Changes affect multiple chart types
└── Risk of breaking other charts
```

### ✅ After (Clean Architecture)
```
chart_xml_writer.dart: ~200 lines (orchestration only)
chart_color_config.dart: ~100 lines (colors)
chart_builders/:
├── chart_style_builder.dart: 15 lines
├── column_bar_chart_builder.dart: 35 lines
├── line_chart_builder.dart: 35 lines
├── area_chart_builder.dart: 30 lines
├── scatter_chart_builder.dart: 40 lines
├── pie_chart_builder.dart: 40 lines
├── radar_chart_builder.dart: 35 lines
└── chart_style_builder_factory.dart: 25 lines
───────────────────────────────────────
Total: ~555 lines (similar size, infinitely better modularity)
```

---

## 🔧 Workflow Example: Generating a Pie Chart

```
1. User calls: sheet.addChart(PieChart(...))
   ↓
2. ChartXmlWriter.generateChartXml(chart)
   ↓
3. Factory.getBuilder(chart) -> Returns PieChartBuilder
   ↓
4. PieChartBuilder.buildProperties(builder, chart) -> Adds slice angle, hole sizes
   ↓
5. PieChartBuilder.buildSeriesStyle(builder, chart, series, index) -> Pulls palette colors from ChartColorConfig
   ↓
6. Complete valid XML generated! ✅
```

---

## 📚 Related Files
- [chart_colors.md](chart_colors.md) — Documented color schemes and strategies.
- [chart_builders/](file:///c:/Users/gonoj/OneDrive/Documentos/GitHub/excel_community/lib/src/utilities/chart_builders/) — Source directory containing builder classes.
- [chart_color_config.dart](file:///c:/Users/gonoj/OneDrive/Documentos/GitHub/excel_community/lib/src/utilities/chart_color_config.dart) — Color configuration source.
