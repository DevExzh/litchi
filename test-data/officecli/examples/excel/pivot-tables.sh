#!/bin/bash
# Pivot Table Showcase — generates pivot-tables.xlsx with 17 pivot tables.
# CLI twin of pivot-tables.py (officecli Python SDK). Both produce an
# equivalent pivot-tables.xlsx. See pivot-tables.md for a per-sheet guide.
#
# Usage: ./pivot-tables.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/pivot-tables.xlsx"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ==========================================================================
# Source data
# ==========================================================================
$CLI set "$FILE" /Sheet1/A1 --prop text=Region
$CLI set "$FILE" /Sheet1/B1 --prop text=Category
$CLI set "$FILE" /Sheet1/C1 --prop text=Product
$CLI set "$FILE" /Sheet1/D1 --prop text=Quarter
$CLI set "$FILE" /Sheet1/E1 --prop text=Sales
$CLI set "$FILE" /Sheet1/F1 --prop text=Quantity
$CLI set "$FILE" /Sheet1/G1 --prop text=Cost
$CLI set "$FILE" /Sheet1/H1 --prop text=Channel
$CLI set "$FILE" /Sheet1/I1 --prop text=Priority
$CLI set "$FILE" /Sheet1/J1 --prop text=Date

# rows: Region Category Product Quarter Sales Quantity Cost Channel Priority Date
write_row() {
  local r="$1"; shift
  local cols=(A B C D E F G H I J)
  local i=0
  for v in "$@"; do
    $CLI set "$FILE" "/Sheet1/${cols[$i]}${r}" --prop text="$v"
    i=$((i + 1))
  done
}

write_row 2  North Electronics Laptop Q1 12500 45 7500 Online High 2025-01-15
write_row 3  North Electronics Phone Q1 8900 120 5340 Retail High 2025-02-10
write_row 4  North Electronics Tablet Q2 6200 38 3720 Online Medium 2025-04-20
write_row 5  North Electronics Laptop Q2 15800 55 9480 Retail High 2025-05-08
write_row 6  North Electronics Phone Q3 11200 150 6720 Online High 2025-07-12
write_row 7  North Electronics Tablet Q4 9500 62 5700 Retail Medium 2025-10-05
write_row 8  North Clothing Jacket Q1 4200 85 2100 Retail Low 2025-01-22
write_row 9  North Clothing Shoes Q2 5600 70 2800 Online Medium 2025-04-15
write_row 10 North Clothing Hat Q3 1800 110 900 Retail Low 2025-08-03
write_row 11 North Clothing Jacket Q4 7800 95 3900 Online High 2025-11-18
write_row 12 North Food Coffee Q1 2400 200 1200 Retail Low 2025-03-01
write_row 13 North Food Snacks Q2 1500 180 750 Online Low 2025-06-10
write_row 14 North Food Juice Q3 1900 160 950 Retail Medium 2025-09-20
write_row 15 North Food Coffee Q4 3200 220 1600 Online Medium 2025-12-01
write_row 16 South Electronics Phone Q1 18500 200 11100 Online High 2024-01-20
write_row 17 South Electronics Laptop Q2 22000 72 13200 Retail High 2024-05-15
write_row 18 South Electronics Tablet Q3 7800 48 4680 Online Medium 2024-08-22
write_row 19 South Electronics Phone Q4 14200 165 8520 Retail High 2024-11-30
write_row 20 South Clothing Shoes Q1 9200 110 4600 Retail Medium 2024-02-14
write_row 21 South Clothing Jacket Q2 6500 78 3250 Online Low 2024-06-01
write_row 22 South Clothing Hat Q3 3100 130 1550 Retail Low 2024-09-10
write_row 23 South Clothing Shoes Q4 8800 98 4400 Online Medium 2024-12-20
write_row 24 South Food Juice Q1 1800 240 900 Retail Low 2024-03-08
write_row 25 South Food Coffee Q2 3500 280 1750 Online Medium 2024-04-25
write_row 26 South Food Snacks Q3 2200 190 1100 Retail Low 2024-07-14
write_row 27 South Food Juice Q4 2800 210 1400 Online Medium 2024-10-18
write_row 28 East Electronics Tablet Q1 5400 35 3240 Online Medium 2025-02-28
write_row 29 East Electronics Laptop Q2 19500 65 11700 Retail High 2025-05-20
write_row 30 East Electronics Phone Q3 13800 180 8280 Online High 2025-08-15
write_row 31 East Electronics Tablet Q4 8200 52 4920 Retail Medium 2025-11-02
write_row 32 East Clothing Hat Q1 2800 140 1400 Retail Low 2025-01-05
write_row 33 East Clothing Jacket Q2 7200 60 3600 Online Medium 2025-06-18
write_row 34 East Clothing Shoes Q3 5500 88 2750 Retail Medium 2025-09-25
write_row 35 East Clothing Hat Q4 3600 105 1800 Online Low 2025-12-10
write_row 36 East Food Snacks Q1 1200 300 600 Retail Low 2025-03-15
write_row 37 East Food Juice Q2 2100 170 1050 Online Medium 2025-04-30
write_row 38 East Food Coffee Q3 2800 230 1400 Retail Medium 2025-07-22
write_row 39 East Food Snacks Q4 1600 250 800 Online Low 2025-10-28
write_row 40 West Electronics Laptop Q1 20500 68 12300 Online High 2024-01-10
write_row 41 West Electronics Phone Q2 16800 190 10080 Retail High 2024-04-05
write_row 42 West Electronics Tablet Q3 8900 55 5340 Online Medium 2024-08-12
write_row 43 West Electronics Laptop Q4 25000 82 15000 Retail High 2024-11-15
write_row 44 West Clothing Jacket Q1 11000 88 5500 Retail Medium 2024-02-22
write_row 45 West Clothing Shoes Q2 7500 95 3750 Online Medium 2024-05-30
write_row 46 West Clothing Hat Q3 4200 120 2100 Retail Low 2024-09-08
write_row 47 West Clothing Jacket Q4 13500 105 6750 Online High 2024-12-01
write_row 48 West Food Coffee Q1 4500 350 2250 Online Medium 2024-03-18
write_row 49 West Food Snacks Q2 2800 280 1400 Online Medium 2024-06-22
write_row 50 West Food Juice Q3 3200 260 1600 Retail Low 2024-07-30
write_row 51 West Food Coffee Q4 5800 400 2900 Online High 2024-10-25

# Chinese-locale source sheet
$CLI add "$FILE" / --type sheet --prop name=CNData
$CLI set "$FILE" /CNData/A1 --prop text=地区
$CLI set "$FILE" /CNData/B1 --prop text=品类
$CLI set "$FILE" /CNData/C1 --prop text=销售额

write_cn() { $CLI set "$FILE" "/CNData/A$1" --prop text="$2"; $CLI set "$FILE" "/CNData/B$1" --prop text="$3"; $CLI set "$FILE" "/CNData/C$1" --prop text="$4"; }
write_cn 2  华东 电子产品 18000
write_cn 3  华东 服装 9500
write_cn 4  华东 食品 4200
write_cn 5  华南 电子产品 22000
write_cn 6  华南 服装 12000
write_cn 7  华南 食品 5800
write_cn 8  华北 电子产品 15000
write_cn 9  华北 服装 7800
write_cn 10 华北 食品 3600
write_cn 11 西南 电子产品 11000
write_cn 12 西南 服装 6500
write_cn 13 西南 食品 2900

# ==========================================================================
# 17 Pivot Tables
# ==========================================================================

# Sheet: 1-Sales Overview
# Features: tabular layout, 2-level rows, column axis, 3 value fields,
#   Cost as percent_of_row, dual page filters, repeat item labels, desc sort
$CLI add "$FILE" / --type sheet --prop name="1-Sales Overview"
$CLI add "$FILE" "/1-Sales Overview" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category \
  --prop cols=Quarter \
  --prop 'values=Sales:sum,Quantity:sum,Cost:sum:percent_of_row' \
  --prop 'filters=Channel,Priority' \
  --prop layout=tabular \
  --prop repeatlabels=true \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop sort=desc \
  --prop name=SalesOverview \
  --prop style=PivotStyleDark2

# Sheet: 2-Market Share
# Features: outline layout, percent_of_col (each region's share per category)
$CLI add "$FILE" / --type sheet --prop name="2-Market Share"
$CLI add "$FILE" "/2-Market Share" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region \
  --prop cols=Category \
  --prop 'values=Sales:sum:percent_of_col' \
  --prop filters=Channel \
  --prop layout=outline \
  --prop grandtotals=both \
  --prop name=MarketShare \
  --prop style=PivotStyleMedium4

# Sheet: 3-Product Deep Dive
# Features: 5 value fields (sum, average, max), no column axis — values
#   become column headers via synthetic "Values" axis, row grand totals only
$CLI add "$FILE" / --type sheet --prop name="3-Product Deep Dive"
$CLI add "$FILE" "/3-Product Deep Dive" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Category,Product \
  --prop 'values=Sales:sum,Sales:average,Sales:max,Quantity:sum,Cost:sum' \
  --prop filters=Region \
  --prop layout=tabular \
  --prop grandtotals=rows \
  --prop subtotals=on \
  --prop sort=desc \
  --prop name=ProductDeepDive \
  --prop style=PivotStyleMedium9

# Sheet: 4-Channel Analysis
# Features: percent_of_total (global share), no filters
$CLI add "$FILE" / --type sheet --prop name="4-Channel Analysis"
$CLI add "$FILE" "/4-Channel Analysis" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Channel \
  --prop cols=Quarter \
  --prop 'values=Sales:sum:percent_of_total,Quantity:sum' \
  --prop layout=outline \
  --prop grandtotals=both \
  --prop name=ChannelTrend \
  --prop style=PivotStyleLight21

# Sheet: 5-Priority Matrix
# Features: blankRows — empty line after each outer group for visual separation
$CLI add "$FILE" / --type sheet --prop name="5-Priority Matrix"
$CLI add "$FILE" "/5-Priority Matrix" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Priority,Region \
  --prop cols=Category \
  --prop 'values=Sales:sum,Cost:sum:percent_of_row' \
  --prop filters=Channel \
  --prop layout=tabular \
  --prop blankrows=true \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop sort=asc \
  --prop name=PriorityMatrix \
  --prop style=PivotStyleDark6

# Sheet: 6-Compact 3-Level
# Features: compact layout — 3-level hierarchy in one indented column
$CLI add "$FILE" / --type sheet --prop name="6-Compact 3-Level"
$CLI add "$FILE" "/6-Compact 3-Level" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category,Product \
  --prop 'values=Sales:sum,Quantity:sum' \
  --prop filters=Priority \
  --prop layout=compact \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop sort=desc \
  --prop name=Compact3Level \
  --prop style=PivotStyleMedium14

# Sheet: 7-No Subtotals
# Features: subtotals=off (flat view), grandtotals=cols (bottom row only),
#   repeatlabels=true (essential when subtotals off — otherwise outer labels vanish)
$CLI add "$FILE" / --type sheet --prop name="7-No Subtotals"
$CLI add "$FILE" "/7-No Subtotals" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category \
  --prop cols=Quarter \
  --prop values=Sales:sum \
  --prop layout=tabular \
  --prop repeatlabels=true \
  --prop grandtotals=cols \
  --prop subtotals=off \
  --prop sort=asc \
  --prop name=FlatView \
  --prop style=PivotStyleLight1

# Sheet: 8-Date Grouping
# Features: automatic date grouping — Date:year creates "2024","2025" buckets,
#   Date:quarter creates "2024-Q1",... sub-buckets. Uses native Excel fieldGroup XML.
$CLI add "$FILE" / --type sheet --prop name="8-Date Grouping"
$CLI add "$FILE" "/8-Date Grouping" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop 'rows=Date:year,Date:quarter' \
  --prop 'values=Sales:sum,Cost:sum' \
  --prop filters=Region \
  --prop layout=outline \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop name=DateGrouping \
  --prop style=PivotStyleMedium7

# Sheet: 9-Top 5 Products
# Features: topN=5 (only top 5 products by first value field), grandtotals=none
$CLI add "$FILE" / --type sheet --prop name="9-Top 5 Products"
$CLI add "$FILE" "/9-Top 5 Products" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Product \
  --prop 'values=Sales:sum,Quantity:sum,Cost:sum' \
  --prop layout=tabular \
  --prop grandtotals=none \
  --prop topN=5 \
  --prop sort=desc \
  --prop name=Top5Products \
  --prop style=PivotStyleDark1

# Sheet: 10-Ultimate
# Features: ALL features combined — tabular + repeatLabels + blankRows +
#   dual filters + 3 mixed-aggregation values + row-only grand totals
$CLI add "$FILE" / --type sheet --prop name="10-Ultimate"
$CLI add "$FILE" "/10-Ultimate" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category \
  --prop cols=Quarter \
  --prop 'values=Sales:sum,Quantity:average,Cost:sum:percent_of_row' \
  --prop 'filters=Channel,Priority' \
  --prop layout=tabular \
  --prop repeatlabels=true \
  --prop blankrows=true \
  --prop grandtotals=rows \
  --prop subtotals=on \
  --prop sort=desc \
  --prop name=UltimatePivot \
  --prop style=PivotStyleDark11

# Sheet: 11-Chinese Locale
# Features: sort=locale (Chinese pinyin: 华北 < 华东 < 华南 < 西南),
#   grandTotalCaption=合计 (custom grand total label)
$CLI add "$FILE" / --type sheet --prop name="11-Chinese Locale"
$CLI add "$FILE" "/11-Chinese Locale" --type pivottable \
  --prop source=CNData!A1:C13 \
  --prop rows=地区,品类 \
  --prop values=销售额:sum \
  --prop layout=tabular \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop sort=locale \
  --prop grandTotalCaption=合计 \
  --prop name=ChineseLocale \
  --prop style=PivotStyleMedium2

# Sheet: 12-Position + Aggregates
# Features: position=D2 (anchor cell override, default is auto-place after source),
#   aggregate=avg (default agg when omitted from a value tuple),
#   value aggregations: count, min, product, countNums (sum/avg/max shown elsewhere)
$CLI add "$FILE" / --type sheet --prop name="12-Position + Aggregates"
$CLI add "$FILE" "/12-Position + Aggregates" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop position=D2 \
  --prop rows=Category \
  --prop 'values=Sales:count,Quantity:min,Quantity:product,Sales:countNums' \
  --prop aggregate=avg \
  --prop layout=tabular \
  --prop grandtotals=both \
  --prop name=PositionAggs \
  --prop style=PivotStyleLight16

# Sheet: 13-Calculated Field
# Features: calculatedField1/2 — user-defined formula fields auto-added as
#   data fields (no need to mention in values=). labelFilter — pre-cache row
#   filter ('Region:beginsWith:N' keeps only Region values starting with N).
$CLI add "$FILE" / --type sheet --prop name="13-Calculated Field"
$CLI add "$FILE" "/13-Calculated Field" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop 'calculatedField1=Margin:=Sales-Cost' \
  --prop 'calculatedField2=Tax:=Sales*0.1' \
  --prop rows=Region \
  --prop values=Sales:sum \
  --prop 'labelFilter=Region:beginsWith:N' \
  --prop layout=tabular \
  --prop grandtotals=both \
  --prop name=CalcField \
  --prop style=PivotStyleMedium3

# Sheet: 14-Statistical
# Features: var / varP (sample + population variance) — completes the aggregate
#   set. showDataAs=running_total as a standalone --prop (vs the tuple form
#   'Field:agg:mode'); applies as default display for all value fields.
$CLI add "$FILE" / --type sheet --prop name="14-Statistical"
$CLI add "$FILE" "/14-Statistical" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region \
  --prop cols=Quarter \
  --prop 'values=Sales:var,Sales:varP,Sales:sum' \
  --prop showDataAs=running_total \
  --prop layout=tabular \
  --prop grandtotals=both \
  --prop name=Statistical \
  --prop style=PivotStyleLight10

# Sheet: 15-Independent Totals
# Features: rowGrandTotals + colGrandTotals as independent toggles
#   (vs the combined grandtotals=both/rows/cols/none), defaultSubtotal=true
#   (default-subtotal flag on every pivotField), sort=locale-desc (reverse
#   pinyin: 西南 > 华南 > 华东 > 华北).
$CLI add "$FILE" / --type sheet --prop name="15-Independent Totals"
$CLI add "$FILE" "/15-Independent Totals" --type pivottable \
  --prop source=CNData!A1:C13 \
  --prop rows=地区 \
  --prop cols=品类 \
  --prop values=销售额:sum \
  --prop rowGrandTotals=true \
  --prop colGrandTotals=false \
  --prop defaultSubtotal=true \
  --prop layout=outline \
  --prop subtotals=on \
  --prop sort=locale-desc \
  --prop name=IndepTotals \
  --prop style=PivotStyleMedium11

# Sheet: 16-Style Flags
# Features: every pivotTableStyleInfo flag wired up — row/col banding,
#   row/col header emphasis, last-column highlight. These map to the five
#   checkboxes in Excel's PivotTable Styles ribbon.
$CLI add "$FILE" / --type sheet --prop name="16-Style Flags"
$CLI add "$FILE" "/16-Style Flags" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category \
  --prop cols=Quarter \
  --prop values=Sales:sum \
  --prop showRowStripes=true \
  --prop showColStripes=true \
  --prop showRowHeaders=true \
  --prop showColHeaders=true \
  --prop showLastColumn=true \
  --prop layout=tabular \
  --prop grandtotals=both \
  --prop name=StyleFlags \
  --prop style=PivotStyleMedium17

# Sheet: 17-Display Toggles
# Features: showDrill=false (hide +/- expand-collapse buttons on every field),
#   mergeLabels=true (merge & center repeated outer-axis item cells —
#   <pivotTableDefinition mergeItem='1'>).
$CLI add "$FILE" / --type sheet --prop name="17-Display Toggles"
$CLI add "$FILE" "/17-Display Toggles" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category \
  --prop values=Sales:sum \
  --prop showDrill=false \
  --prop mergeLabels=true \
  --prop layout=outline \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop name=DisplayToggles \
  --prop style=PivotStyleLight19

$CLI close "$FILE"

$CLI validate "$FILE"
echo "Generated: $FILE"
echo "  19 sheets (Sheet1 + CNData + 17 pivot tables)"
