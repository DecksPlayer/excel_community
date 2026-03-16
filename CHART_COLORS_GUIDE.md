# 🎨 Color Strategy by Chart Type

## Executive Summary

A comprehensive color system has been implemented for all chart types in the `excel_community` library, ensuring professional, distinguishable, and aesthetically pleasing visualizations.

---

## 📊 Series Charts (Column, Bar, Line, Area, Scatter)

### Color Palette (12 colors)
```
1.  #4472C4 - Blue
2.  #ED7D31 - Orange
3.  #70AD47 - Green
4.  #FFC000 - Gold
5.  #5B9BD5 - Light Blue
6.  #C5504B - Red
7.  #8064A2 - Purple
8.  #4BACC6 - Cyan
9.  #9BBB59 - Olive
10. #F79646 - Light Orange
11. #17B897 - Teal
12. #E83352 - Crimson
```

### Automatic Rotation
If there are more than 12 series, colors are reused using `i % 12`.

---

## 1️⃣ Column Chart

**Features:**
- ✅ Solid colors (100% opacity)
- ✅ Borders of the same color as the fill
- ✅ Border thickness: 9525 EMUs (thin)

**Usage:**
```dart
ColumnChart(
  title: "Sales by Quarter",
  series: series,  // Each series will have a different color
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Vertical bars with distinguishable solid colors
- Ideal for comparing categories

---

## 2️⃣ Bar Chart (Horizontal Bar Chart)

**Features:**
- ✅ Identical to Column Chart but horizontal
- ✅ Solid colors (100% opacity)
- ✅ Borders of the same color

**Usage:**
```dart
BarChart(
  title: "Revenue by Product",
  series: series,
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Horizontal bars with solid colors
- Better for long labels

---

## 3️⃣ Line Chart

**Features:**
- ✅ Thick lines: 28575 EMUs
- ✅ Circular markers (size 5)
- ✅ Markers of the same color as the line
- ✅ No transparency (100% opacity)

**Usage:**
```dart
LineChart(
  title: "Trends Over Time",
  series: series,
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Colorful lines with circular points
- Ideal for temporal trends

---

## 4️⃣ Area Chart

**Features:**
- ✅ Lines: 90% opacity (alpha: 90000)
- ✅ Fill: 50% opacity (alpha: 50000)
- ✅ Line thickness: 28575 EMUs (thick)

**Usage:**
```dart
AreaChart(
  title: "Market Share Evolution",
  series: series,
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Semi-transparent areas that allow seeing overlaps
- More opaque lines to define contours
- Perfect for showing accumulation or parts of a whole

---

## 5️⃣ Scatter Chart

**Features:**
- ✅ Circular points (size 7)
- ✅ Solid fill of the series color
- ✅ White border (#FFFFFF) thickness 9525
- ✅ No transparency

**Usage:**
```dart
ScatterChart(
  title: "Correlation Analysis",
  series: series,
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Colorful dots with a white halo
- The white border helps distinguish overlapping points
- Ideal for correlations and distributions

---

## 🥧 Circular Charts (Pie, Doughnut)

### Color Palette (20 colors)
```
1.  #4472C4 - Blue          11. #255E91 - Navy Blue
2.  #ED7D31 - Orange        12. #43682B - Dark Green
3.  #A5A5A5 - Gray          13. #C5504B - Red
4.  #FFC000 - Gold          14. #8064A2 - Purple
5.  #5B9BD5 - Light Blue    15. #4BACC6 - Cyan
6.  #70AD47 - Green         16. #F79646 - Light Orange
7.  #264478 - Dark Blue     17. #9BBB59 - Olive
8.  #9E480E - Brown         18. #E83352 - Crimson
9.  #636363 - Dark Gray     19. #17B897 - Teal
10. #997300 - Dark Gold     20. #FF6F61 - Coral
```

### Assignment Algorithm
1. Shuffle (randomize) the palette
2. Take the first N colors (N = number of segments)
3. Assign one by one to each segment

**Result:** Each execution generates random combinations WITHOUT repetition

---

## 6️⃣ Pie Chart

**Features:**
- ✅ Random colors without repetition
- ✅ Solid colors (100% opacity)
- ✅ 20 colors available
- ✅ Shuffle before assigning

**Usage:**
```dart
PieChart(
  title: "Market Share",
  series: [series[0]], // Only one series
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Each segment has a different vibrant color
- Randomized colors every time
- Maximum 20 segments with unique colors

---

## 7️⃣ Doughnut Chart

**Features:**
- ✅ Identical to Pie Chart in colors
- ✅ Central hole (holeSize: 50%)
- ✅ Random colors without repetition

**Usage:**
```dart
DoughnutChart(
  title: "Budget Distribution",
  series: [series[0]], // Only one series
  anchor: anchor,
  showLegend: true,
)
```

**Visual:**
- Like Pie Chart but with a central hole
- Same randomized color system

---

## 🕸️ Radar Chart

### Color Palette (8 colors)
```
1. #4472C4 - Blue
2. #ED7D31 - Orange
3. #70AD47 - Green
4. #FFC000 - Gold
5. #5B9BD5 - Light Blue
6. #C5504B - Red
7. #8064A2 - Purple
8. #4BACC6 - Cyan
```

---

## 8️⃣ Radar Chart - Filled

**Features:**
- ✅ Lines: 85% opacity (alpha: 85000)
- ✅ Fill: 45% opacity (alpha: 45000)
- ✅ Line thickness: 28575 EMUs (thick)

**Usage:**
```dart
RadarChart(
  title: "Skills Assessment",
  series: series,
  anchor: anchor,
  showLegend: true,
  filled: true, // ← IMPORTANT
)
```

**Visual:**
- Very transparent areas (45%) allow seeing overlaps
- More visible lines (85%)
- Perfect for comparing multiple profiles

---

## 9️⃣ Radar Chart - Lines

**Features:**
- ✅ Lines: 85% opacity (alpha: 85000)
- ✅ No fill
- ✅ Line thickness: 28575 EMUs

**Usage:**
```dart
RadarChart(
  title: "Performance Metrics",
  series: series,
  anchor: anchor,
  showLegend: true,
  filled: false, // ← IMPORTANT
)
```

**Visual:**
- Only contours without fill
- Cleaner when there are many overlapping series

---

## 📊 Transparency Comparison

| Chart Type | Lines | Fill | Markers | Borders |
|------------|-------|------|---------|---------|
| Column/Bar | - | 100% | - | 100% |
| Line | 100% | - | 100% | - |
| Area | 90% | 50% | - | - |
| Scatter | - | 100% | 100% | White |
| Pie/Doughnut | - | 100% | - | - |
| Radar (filled) | 85% | 45% | - | - |
| Radar (lines) | 85% | - | - | - |

---

## 🎯 Applied Design Principles

### 1. **Distinguishability**
- Colors are sufficiently different from each other
- Avoids visual confusion

### 2. **Strategic Transparency**
- **Area/Radar charts:** Transparency to see overlaps
- **Bar/Scatter charts:** Solid for maximum clarity

### 3. **Consistency**
- Same base palette for all series charts
- Extended palette for circular charts (more segments)

### 4. **Professionalism**
- Colors based on Office/Excel palettes
- Neither too bright nor too dull

### 5. **Accessibility**
- Includes tone variations (light/dark)
- Good contrasts

---

## 📝 Implementation Notes

### Measurement Units
- **EMUs (English Metric Units):** 914400 EMUs = 1 inch
- **Thin line thickness:** 9525 EMUs ≈ 0.75 pt
- **Thick line thickness:** 28575 EMUs ≈ 2.25 pt

### Alpha Values (Transparency)
- **100% opaque:** No `<a:alpha>` element
- **90% opaque:** `alpha="90000"`
- **85% opaque:** `alpha="85000"`
- **50% opaque:** `alpha="50000"`
- **45% opaque:** `alpha="45000"`

The alpha value ranges from 0 to 100000 (0% to 100%)

---

## 🧪 Generated Test Files

Run `dart run test_all_colors.dart` to generate:

1. `COLOR_TEST_column_chart.xlsx`
2. `COLOR_TEST_bar_chart.xlsx`
3. `COLOR_TEST_line_chart.xlsx`
4. `COLOR_TEST_area_chart.xlsx`
5. `COLOR_TEST_scatter_chart.xlsx`
6. `COLOR_TEST_pie_chart.xlsx`
7. `COLOR_TEST_doughnut_chart.xlsx`
8. `COLOR_TEST_radar_filled_chart.xlsx`
9. `COLOR_TEST_radar_lines_chart.xlsx`

**All** charts include:
- 3 data series (except circular which use 1)
- 6 categories
- Legend enabled
- Descriptive titles

---

## ✅ Quality Verification

Open the generated files in Microsoft Excel and verify:

- ✅ All colors are distinct and visible
- ✅ Transparencies work correctly
- ✅ Borders and markers look good
- ✅ Legends show correct colors
- ✅ No repeated colors in the same chart
- ✅ Circular charts have random variety

---

## 🚀 Recommended Usage

### Column/Bar Charts
- Comparisons between categories
- Discrete data
- Each category clearly separated

### Line Charts
- Temporal trends
- Continuous data
- Changes over time

### Area Charts
- Accumulation of values
- Contribution of parts to the total
- When overlap is important

### Scatter Charts
- Correlations
- Distributions
- Relationships between variables

### Pie/Doughnut Charts
- Parts of a whole (percentages)
- Maximum 6-8 segments for readability
- When the total adds up to 100%

### Radar Charts
- Compare multiple metrics
- Multidimensional profiles
- Skills assessments

---

## 📦 Conclusion

Each chart type has a color strategy optimized for its specific use case:

- **Bar/Column charts:** Solid and distinguishable
- **Line charts:** Visible markers
- **Area/Radar charts:** Transparent to see overlaps
- **Circular charts:** Random for variety
- **Scatter charts:** White borders for separation

This system ensures professional and easy-to-interpret visualizations across all cases.

---

**Implementation Date:** March 14, 2026  
**Library:** excel_community  
**Implementation File:** `lib/src/utilities/chart_xml_writer.dart`
