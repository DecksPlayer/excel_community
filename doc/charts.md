# 📊 Working with Charts

Excel Community allows you to insert standard charts directly into worksheets. The chart styles are designed using professional, modern color palettes, and the emitted XML is fully compliant with the OOXML standard for compatibility with Microsoft Excel, Excel 365, and Google Sheets.

---

## 📈 Supported Chart Types

- 📶 **ColumnChart** - Vertical clustered bar charts.
- ▤ **BarChart** - Horizontal clustered bar charts.
- 📈 **LineChart** - Line charts with circular data markers.
- 🏞️ **AreaChart** - Area charts with translucent fills.
- 🍰 **PieChart** - standard pie charts (supports single data series).
- 🍩 **DoughnutChart** - Pie charts with a center hole (supports single data series).
- ⁛ **ScatterChart** - Scatter / XY plots (requires numeric X and Y values).
- 🕸️ **RadarChart** - Radar / spider charts (supports line-only or filled variants).

---

## 📝 Basic Chart Example

Adding a chart involves populating the sheet cells with data, defining a series with range strings (using absolute cell address format), and using `sheet.addChart()` to place the chart.

```dart
import 'package:excel_community/excel_community.dart';

void createChart() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // 1. Add headers
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Month"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Sales"));
  sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Expenses"));

  // 2. Add data rows
  final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  final sales = [45, 55, 42, 60, 58, 65];
  final expenses = [35, 48, 52, 45, 62, 55];

  for (var i = 0; i < months.length; i++) {
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(months[i]));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), IntCellValue(sales[i]));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), IntCellValue(expenses[i]));
  }

  // 3. Create the Chart definition
  var chart = ColumnChart(
    title: "Monthly Sales vs Expenses",
    series: [
      ChartSeries(
        name: "Sales",
        categoriesRange: r"Sheet1!$A$2:$A$7", // X-Axis Labels
        valuesRange: r"Sheet1!$B$2:$B$7",     // Y-Axis Values
      ),
      ChartSeries(
        name: "Expenses",
        categoriesRange: r"Sheet1!$A$2:$A$7",
        valuesRange: r"Sheet1!$C$2:$C$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 4, row: 1, width: 12, height: 16), // Position at column E, row 2
    showLegend: true,
  );

  // 4. Attach chart to sheet
  sheet.addChart(chart);

  // 5. Save workbook
  var bytes = excel.save();
}
```

---

## 🔍 Specific Chart Options

### 1. Column & Bar Charts
Clustered bars with solid fills and thin borders matching the fill color.

```dart
var chart = ColumnChart(
  title: "Quarterly Revenue",
  series: [
    ChartSeries(
      name: "Q1",
      categoriesRange: r"Sheet1!$A$2:$A$5",
      valuesRange: r"Sheet1!$B$2:$B$5",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1),
);
```

### 2. Line Charts
Line graphs displaying trends. Automatically draws thick colored lines with circular markers.

```dart
var chart = LineChart(
  title: "Website Traffic Trends",
  series: [
    ChartSeries(
      name: "Visitors",
      categoriesRange: r"Sheet1!$A$2:$A$13",
      valuesRange: r"Sheet1!$B$2:$B$13",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1),
);
```

### 3. Area Charts
Area plots featuring translucent styling (fills set to 50% opacity and lines set to 90% opacity).

```dart
var chart = AreaChart(
  title: "Growth Comparison",
  series: [
    ChartSeries(
      name: "Product A",
      categoriesRange: r"Sheet1!$A$2:$A$7",
      valuesRange: r"Sheet1!$B$2:$B$7",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1),
);
```

### 4. Pie & Doughnut Charts
Proportional graphs that require a single data series. Slices are colored using randomized colors from the palette.

```dart
var chart = PieChart( // Or DoughnutChart()
  title: "Market Share",
  series: [
    ChartSeries(
      name: "Shares",
      categoriesRange: r"Sheet1!$A$2:$A$6",
      valuesRange: r"Sheet1!$B$2:$B$6",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  showLegend: true,
);
```

### 5. Scatter Charts (XY)
Used to show relationships between numerical variables. Both ranges must point to cells containing numbers (`IntCellValue` or `DoubleCellValue`).

```dart
var chart = ScatterChart(
  title: "Height vs Weight Correlation",
  series: [
    ChartSeries(
      name: "Subjects",
      categoriesRange: r"Sheet1!$A$2:$A$20",  // Numeric X values
      valuesRange: r"Sheet1!$B$2:$B$20",      // Numeric Y values
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1),
);
```

### 6. Radar Charts
Spider charts that support line-only or filled segments.

```dart
var chart = RadarChart(
  title: "Performance Metrics",
  series: [...],
  anchor: ChartAnchor.at(column: 5, row: 1),
  filled: true, // Set to false for line-only
);
```

---

## ⚓ Positioning Charts (`ChartAnchor`)

A `ChartAnchor` defines the location and dimensions of a chart.

```dart
// Auto-sized positioning
ChartAnchor.at(
  column: 5,      // 0-indexed starting column (Col F)
  row: 1,         // 0-indexed starting row (Row 2)
  width: 10,      // (Optional) Width spanning 10 columns
  height: 15,     // (Optional) Height spanning 15 rows
)

// Coordinate-bounded positioning
ChartAnchor(
  fromColumn: 5,
  fromRow: 1,
  toColumn: 15,   // Ends at Column P
  toRow: 16,      // Ends at Row 17
)
```

---

## 📑 Multiple Charts

You can attach multiple charts to the same worksheet, or split them across different worksheets:

```dart
var sheet1 = excel['Sales'];
var sheet2 = excel['Metrics'];

sheet1.addChart(ColumnChart(
  title: "Monthly Sales",
  series: [...],
  anchor: ChartAnchor.at(column: 5, row: 1),
));

// Stack another chart below the first one
sheet1.addChart(LineChart(
  title: "Yearly Trend",
  series: [...],
  anchor: ChartAnchor.at(column: 5, row: 18),
));
```

---

## 🌈 Chart Colors

Charts automatically cycle through a curated 12-color palette. If your chart includes more than 12 series, the colors repeat.

For a detailed color strategy guide and exact HEX color values, see the [Chart Colors Guide](chart_colors.md).

---

## ⚠️ Compatibility Notes

- **Range Reference Format**: Always specify ranges using absolute Excel notation, including the sheet name (e.g. `r"Sheet1!$A$2:$A$10"`). Using a raw string literal prefixed with `r` prevents Dart from treating `$` as a string interpolation token.
- **Series Names**: All series names are formatted to comply with Excel’s strict requirements so that files open cleanly without parsing errors.
