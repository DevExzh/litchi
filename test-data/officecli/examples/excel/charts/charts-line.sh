#!/bin/bash
# Line Charts Showcase — generates charts-line.xlsx exercising every line chart
# feature officecli supports: line, lineStacked, linePercentStacked, line3d, with
# markers, smoothing, dash patterns, axis scaling, gridlines, data labels, legend
# layout, reference lines, secondary axis, gradients, shadows, manual layout,
# data tables, drop/hi-low lines, up-down bars, and 3D rotation.
#
# CLI twin of charts-line.py (officecli Python SDK). Both produce an equivalent
# charts-line.xlsx.
#
# 9 sheets (Sheet1 data + 8 chart sheets), 32 charts total.
#
# Usage:
#   ./charts-line.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-line.xlsx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Source data — shared across all charts
# ==========================================================================
echo "--- Populating source data ---"
officecli set "$FILE" /Sheet1/A1 --prop text=Month --prop bold=true
officecli set "$FILE" /Sheet1/B1 --prop text=East --prop bold=true
officecli set "$FILE" /Sheet1/C1 --prop text=South --prop bold=true
officecli set "$FILE" /Sheet1/D1 --prop text=North --prop bold=true
officecli set "$FILE" /Sheet1/E1 --prop text=West --prop bold=true

officecli set "$FILE" /Sheet1/A2 --prop text=Jan; officecli set "$FILE" /Sheet1/B2 --prop text=120; officecli set "$FILE" /Sheet1/C2 --prop text=95;  officecli set "$FILE" /Sheet1/D2 --prop text=88;  officecli set "$FILE" /Sheet1/E2 --prop text=110
officecli set "$FILE" /Sheet1/A3 --prop text=Feb; officecli set "$FILE" /Sheet1/B3 --prop text=135; officecli set "$FILE" /Sheet1/C3 --prop text=108; officecli set "$FILE" /Sheet1/D3 --prop text=92;  officecli set "$FILE" /Sheet1/E3 --prop text=118
officecli set "$FILE" /Sheet1/A4 --prop text=Mar; officecli set "$FILE" /Sheet1/B4 --prop text=148; officecli set "$FILE" /Sheet1/C4 --prop text=115; officecli set "$FILE" /Sheet1/D4 --prop text=105; officecli set "$FILE" /Sheet1/E4 --prop text=130
officecli set "$FILE" /Sheet1/A5 --prop text=Apr; officecli set "$FILE" /Sheet1/B5 --prop text=162; officecli set "$FILE" /Sheet1/C5 --prop text=128; officecli set "$FILE" /Sheet1/D5 --prop text=118; officecli set "$FILE" /Sheet1/E5 --prop text=145
officecli set "$FILE" /Sheet1/A6 --prop text=May; officecli set "$FILE" /Sheet1/B6 --prop text=155; officecli set "$FILE" /Sheet1/C6 --prop text=142; officecli set "$FILE" /Sheet1/D6 --prop text=125; officecli set "$FILE" /Sheet1/E6 --prop text=138
officecli set "$FILE" /Sheet1/A7 --prop text=Jun; officecli set "$FILE" /Sheet1/B7 --prop text=178; officecli set "$FILE" /Sheet1/C7 --prop text=155; officecli set "$FILE" /Sheet1/D7 --prop text=138; officecli set "$FILE" /Sheet1/E7 --prop text=162
officecli set "$FILE" /Sheet1/A8 --prop text=Jul; officecli set "$FILE" /Sheet1/B8 --prop text=195; officecli set "$FILE" /Sheet1/C8 --prop text=168; officecli set "$FILE" /Sheet1/D8 --prop text=145; officecli set "$FILE" /Sheet1/E8 --prop text=175
officecli set "$FILE" /Sheet1/A9 --prop text=Aug; officecli set "$FILE" /Sheet1/B9 --prop text=210; officecli set "$FILE" /Sheet1/C9 --prop text=175; officecli set "$FILE" /Sheet1/D9 --prop text=152; officecli set "$FILE" /Sheet1/E9 --prop text=190
officecli set "$FILE" /Sheet1/A10 --prop text=Sep; officecli set "$FILE" /Sheet1/B10 --prop text=188; officecli set "$FILE" /Sheet1/C10 --prop text=160; officecli set "$FILE" /Sheet1/D10 --prop text=140; officecli set "$FILE" /Sheet1/E10 --prop text=170
officecli set "$FILE" /Sheet1/A11 --prop text=Oct; officecli set "$FILE" /Sheet1/B11 --prop text=172; officecli set "$FILE" /Sheet1/C11 --prop text=148; officecli set "$FILE" /Sheet1/D11 --prop text=130; officecli set "$FILE" /Sheet1/E11 --prop text=155
officecli set "$FILE" /Sheet1/A12 --prop text=Nov; officecli set "$FILE" /Sheet1/B12 --prop text=165; officecli set "$FILE" /Sheet1/C12 --prop text=135; officecli set "$FILE" /Sheet1/D12 --prop text=122; officecli set "$FILE" /Sheet1/E12 --prop text=148
officecli set "$FILE" /Sheet1/A13 --prop text=Dec; officecli set "$FILE" /Sheet1/B13 --prop text=198; officecli set "$FILE" /Sheet1/C13 --prop text=158; officecli set "$FILE" /Sheet1/D13 --prop text=142; officecli set "$FILE" /Sheet1/E13 --prop text=180

# ==========================================================================
# Sheet: 1-Line Fundamentals
# ==========================================================================
echo "--- 1-Line Fundamentals ---"
officecli add "$FILE" / --type sheet --prop name="1-Line Fundamentals"

# Chart 1: Basic line with inline named series and categories
# Features: chartType=line, inline series (series1=Name:v1,v2,...),
#   categories, colors, catTitle, axisTitle, axisfont, gridlines
officecli add "$FILE" "/1-Line Fundamentals" --type chart \
  --prop chartType=line \
  --prop title="Quarterly Revenue" \
  --prop series1="Product A:120,180,210,250" \
  --prop series2="Product B:90,140,160,200" \
  --prop series3="Product C:60,85,110,145" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop catTitle=Quarter --prop axisTitle=Revenue \
  --prop axisfont=9:C00000:Arial \
  --prop gridlines=D9D9D9:0.5:dot

# Chart 2: Line with cell-range series (dotted syntax) and markers
# Features: series.name/values/categories (cell range via dotted syntax),
#   showMarkers, marker (style:size:color), minorGridlines
officecli add "$FILE" "/1-Line Fundamentals" --type chart \
  --prop chartType=line \
  --prop title="East Region Trend" \
  --prop series1.name=East \
  --prop series1.values=Sheet1!B2:B13 \
  --prop series1.categories=Sheet1!A2:A13 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop showMarkers=true --prop marker=circle:6:2E75B6 \
  --prop gridlines=D9D9D9:0.5:dot \
  --prop minorGridlines=EEEEEE:0.3:dot

# Chart 3: Line from dataRange with all four regions
# Features: dataRange (auto-reads headers as series names), marker=diamond,
#   lineWidth, legend=bottom, legendfont
officecli add "$FILE" "/1-Line Fundamentals" --type chart \
  --prop chartType=line \
  --prop title="All Regions — Full Year" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=2E75B6,70AD47,FFC000,C00000 \
  --prop showMarkers=true --prop marker=diamond:5:333333 \
  --prop lineWidth=2 \
  --prop legend=bottom \
  --prop legendfont=9:58626E:Calibri

# Chart 4: Line with inline data shorthand and marker=none
# Features: data (inline shorthand Name:v1;Name2:v2), marker=none, legend=right
officecli add "$FILE" "/1-Line Fundamentals" --type chart \
  --prop chartType=line \
  --prop title="Simple Two-Series" \
  --prop "data=Actual:80,120,160,200,240;Target:100,130,160,190,220" \
  --prop categories="Week 1,Week 2,Week 3,Week 4,Week 5" \
  --prop colors=0070C0,FF0000 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop marker=none \
  --prop legend=right

# ==========================================================================
# Sheet: 2-Line Styles
# ==========================================================================
echo "--- 2-Line Styles ---"
officecli add "$FILE" / --type sheet --prop name="2-Line Styles"

# Chart 1: Smooth line with thick width and shadow
# Features: smooth=true (Bezier curves), lineWidth=2.5, gridlines=none,
#   axisVisible=false (hide both axes for sparkline-like minimal look),
#   series.shadow (color-blur-angle-dist-opacity)
officecli add "$FILE" "/2-Line Styles" --type chart \
  --prop chartType=line \
  --prop title="Smooth Curves with Shadow" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop smooth=true --prop lineWidth=2.5 \
  --prop colors=0070C0,00B050,FFC000,FF0000 \
  --prop gridlines=none \
  --prop axisVisible=false \
  --prop series.shadow=000000-4-315-2-40

# Chart 2: Dashed lines — all dash styles demonstrated
# Note: lineDash applies to ALL series. Supported values:
#   solid, dot, dash, dashdot, longdash, longdashdot, longdashdotdot
# Features: lineDash (applied globally to all series), lineWidth
officecli add "$FILE" "/2-Line Styles" --type chart \
  --prop chartType=line \
  --prop title="Dash Pattern Gallery" \
  --prop series1="solid:120,135,148,162,155" \
  --prop series2="dot:95,108,115,128,142" \
  --prop series3="dash:88,92,105,118,125" \
  --prop series4="dashdot:110,118,130,145,138" \
  --prop categories=Jan,Feb,Mar,Apr,May \
  --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop lineDash=dash --prop lineWidth=2 \
  --prop legend=bottom

# Chart 3: Multiple marker styles — circle, square, triangle, star
# Note: marker applies to ALL series. Supported styles:
#   circle, diamond, square, triangle, star, x, plus, dash, dot, none
# Features: marker=square:7:color (style:size:fillColor),
#   series.outline (white border around markers/lines)
officecli add "$FILE" "/2-Line Styles" --type chart \
  --prop chartType=line \
  --prop title="Marker Style Showcase" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop showMarkers=true --prop marker=square:7:4472C4 \
  --prop lineWidth=1.5 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop series.outline=FFFFFF-0.5

# Chart 4: Transparent lines with gradient plot area and styled title
# Features: transparency=30 (30% transparent), plotFill gradient,
#   chartFill, title.font/size/color/bold, roundedCorners
officecli add "$FILE" "/2-Line Styles" --type chart \
  --prop chartType=line \
  --prop title="Translucent Lines on Gradient" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop lineWidth=3 --prop smooth=true \
  --prop transparency=30 \
  --prop plotFill=F0F4F8-D6E4F0:90 \
  --prop chartFill=FFFFFF \
  --prop colors=1F4E79,2E75B6,5B9BD5,9DC3E6 \
  --prop title.font=Georgia --prop title.size=14 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop roundedCorners=true

# ==========================================================================
# Sheet: 3-Line Variants
# ==========================================================================
echo "--- 3-Line Variants ---"
officecli add "$FILE" / --type sheet --prop name="3-Line Variants"

# Chart 1: Stacked line chart
# Features: lineStacked (cumulative stacking), majorTickMark=outside, tickLabelPos=low
officecli add "$FILE" "/3-Line Variants" --type chart \
  --prop chartType=lineStacked \
  --prop title="Cumulative Sales by Region" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop catTitle=Month --prop axisTitle=Cumulative \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop majorTickMark=outside --prop tickLabelPos=low

# Chart 2: 100% stacked line chart with axis number format
# Features: linePercentStacked (each month sums to 100%),
#   axisNumFmt (value axis number format)
officecli add "$FILE" "/3-Line Variants" --type chart \
  --prop chartType=linePercentStacked \
  --prop title="Regional Contribution %" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=1F4E79,2E75B6,9DC3E6,BDD7EE \
  --prop axisNumFmt=0% \
  --prop legend=right \
  --prop gridlines=E0E0E0:0.5:solid

# Chart 3: 3D line chart with perspective
# Features: line3d (3D line chart), view3d (rotX,rotY,perspective),
#   style/styleId (preset chart style 1-48)
officecli add "$FILE" "/3-Line Variants" --type chart \
  --prop chartType=line3d \
  --prop title="3D Regional Trends" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=15,20,30 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop chartFill=F8F8F8 \
  --prop style=3

# Chart 4: Stacked line with area fill and data table
# Features: dataTable=true (show value table below chart),
#   legend=none (hidden because data table shows series names)
officecli add "$FILE" "/3-Line Variants" --type chart \
  --prop chartType=lineStacked \
  --prop title="Stacked with Data Table" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop dataTable=true \
  --prop legend=none \
  --prop lineWidth=1.5 \
  --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
  --prop plotFill=FAFAFA

# ==========================================================================
# Sheet: 4-Axis & Gridlines
# ==========================================================================
echo "--- 4-Axis & Gridlines ---"
officecli add "$FILE" / --type sheet --prop name="4-Axis & Gridlines"

# Chart 1: Custom axis scaling — min, max, majorUnit
# Features: axisMin, axisMax, majorUnit, minorUnit,
#   axisLine (value axis line styling — red), catAxisLine (category axis line — blue)
officecli add "$FILE" "/4-Axis & Gridlines" --type chart \
  --prop chartType=line \
  --prop title="Custom Axis Scale (80–220)" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop axisMin=80 --prop axisMax=220 --prop majorUnit=20 \
  --prop minorUnit=10 \
  --prop showMarkers=true --prop marker=circle:4:4472C4 \
  --prop gridlines=D0D0D0:0.5:solid \
  --prop minorGridlines=EEEEEE:0.3:dot \
  --prop axisLine=C00000:1.5:solid \
  --prop catAxisLine=2E75B6:1.5:solid

# Demonstrate chart-axis element (path: /SheetName/chart[N]/axis[@role=ROLE]).
# Properties: min, max, format, majorGridlines, labelRotation.
# These are the same semantics as axisMin/axisMax/gridlines/labelrotation at
# chart level but applied through the dedicated sub-element path, which also
# exposes role, dispUnits, majorUnit, title, visible, logBase.
officecli set "$FILE" "/4-Axis & Gridlines/chart[1]/axis[@role=value]" \
  --prop min=80 --prop max=220 \
  --prop format="#,##0" \
  --prop majorGridlines=true \
  --prop labelRotation=0

# Chart 2: Logarithmic scale with display units
# Features: logBase=10 (logarithmic scale), marker=triangle
officecli add "$FILE" "/4-Axis & Gridlines" --type chart \
  --prop chartType=line \
  --prop title="Exponential Growth (Log Scale)" \
  --prop series1="Growth:1,5,25,125,625,3125" \
  --prop categories="Year 1,Year 2,Year 3,Year 4,Year 5,Year 6" \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop logBase=10 \
  --prop colors=C00000 \
  --prop lineWidth=2.5 \
  --prop showMarkers=true --prop marker=triangle:7:C00000 \
  --prop axisTitle="Value (log)" \
  --prop catTitle=Year \
  --prop gridlines=E0E0E0:0.5:dash

# Chart 3: Reversed axis and hidden axes
# Features: axisReverse=true (value axis direction flipped), smooth + markers together
officecli add "$FILE" "/4-Axis & Gridlines" --type chart \
  --prop chartType=line \
  --prop title="Reversed Value Axis" \
  --prop series1="Depth:0,50,120,200,350,500" \
  --prop categories="Station A,Station B,Station C,Station D,Station E,Station F" \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop axisReverse=true \
  --prop colors=0070C0 \
  --prop lineWidth=2 \
  --prop showMarkers=true --prop marker=diamond:6:0070C0 \
  --prop smooth=true \
  --prop axisTitle="Depth (m)" \
  --prop gridlines=D9D9D9:0.5:solid

# Chart 4: Display units and tick mark styles
# Features: dispUnits=thousands (display units label),
#   majorTickMark=outside, minorTickMark=inside, marker=star
officecli add "$FILE" "/4-Axis & Gridlines" --type chart \
  --prop chartType=line \
  --prop title="Revenue (in Thousands)" \
  --prop series1="Revenue:12000,18500,22000,31000,45000,52000" \
  --prop series2="Cost:8000,11000,14000,19500,28000,33000" \
  --prop categories=2020,2021,2022,2023,2024,2025 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop dispUnits=thousands \
  --prop colors=2E75B6,C00000 \
  --prop lineWidth=2 \
  --prop majorTickMark=outside --prop minorTickMark=inside \
  --prop showMarkers=true --prop marker=star:7:2E75B6 \
  --prop catTitle=Year --prop axisTitle="Amount (K)"

# ==========================================================================
# Sheet: 5-Labels & Legend
# ==========================================================================
echo "--- 5-Labels & Legend ---"
officecli add "$FILE" / --type sheet --prop name="5-Labels & Legend"

# Chart 1: Data labels at various positions with number format
# Features: dataLabels=true, labelPos=top, labelFont (size:color:bold),
#   dataLabels.numFmt (number format), dataLabels.separator
officecli add "$FILE" "/5-Labels & Legend" --type chart \
  --prop chartType=line \
  --prop title="Sales with Labels" \
  --prop series1="Revenue:120,180,210,250,280" \
  --prop categories=Jan,Feb,Mar,Apr,May \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=4472C4 \
  --prop lineWidth=2 \
  --prop showMarkers=true --prop marker=circle:6:4472C4 \
  --prop dataLabels=true --prop labelPos=top \
  --prop labelFont=9:333333:true \
  --prop dataLabels.numFmt=#,##0 \
  --prop "dataLabels.separator=: "

# Chart 2: Custom individual data labels (highlight peak)
# Features: dataLabel{N}.delete (hide specific labels),
#   dataLabel{N}.text (custom text on specific point),
#   point{N}.color (highlight individual data point marker in red),
#   dataLabel{N}.y (manual vertical position of individual label, 0-1 fraction)
officecli add "$FILE" "/5-Labels & Legend" --type chart \
  --prop chartType=line \
  --prop title="Peak Highlight" \
  --prop series1="Sales:88,120,165,210,195,178" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=2E75B6 \
  --prop lineWidth=2.5 --prop smooth=true \
  --prop showMarkers=true --prop marker=circle:5:2E75B6 \
  --prop dataLabels=true --prop labelPos=top \
  --prop dataLabel1.delete=true --prop dataLabel2.delete=true \
  --prop point4.color=C00000 \
  --prop dataLabel4.text="Peak: 210" \
  --prop dataLabel4.y=0.15 \
  --prop dataLabel5.delete=true --prop dataLabel6.delete=true

# Chart 3: Legend positioning and overlay
# Features: legend=top, legend.overlay=true (legend overlays chart area),
#   legendfont (size:color:fontname)
officecli add "$FILE" "/5-Labels & Legend" --type chart \
  --prop chartType=line \
  --prop title="Legend Overlay on Chart" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop lineWidth=2 \
  --prop legend=top \
  --prop legend.overlay=true \
  --prop legendfont=10:1F4E79:Calibri \
  --prop plotFill=F5F5F5

# Chart 4: Manual layout — plotArea, title, and legend positioning
# Features: plotArea.x/y/w/h (plot area manual layout, 0-1 fraction),
#   title.x/y (title position), legend.x/y/w/h (legend position/size)
officecli add "$FILE" "/5-Labels & Legend" --type chart \
  --prop chartType=line \
  --prop title="Manual Layout Control" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
  --prop lineWidth=1.5 \
  --prop plotArea.x=0.12 --prop plotArea.y=0.18 \
  --prop plotArea.w=0.82 --prop plotArea.h=0.55 \
  --prop title.x=0.25 --prop title.y=0.02 \
  --prop legend.x=0.15 --prop legend.y=0.82 \
  --prop legend.w=0.7 --prop legend.h=0.12 \
  --prop title.font=Arial --prop title.size=13 \
  --prop title.bold=true

# ==========================================================================
# Sheet: 6-Effects & Advanced
# ==========================================================================
echo "--- 6-Effects & Advanced ---"
officecli add "$FILE" / --type sheet --prop name="6-Effects & Advanced"

# Chart 1: Secondary axis — two series on different scales
# Features: secondaryAxis=2 (series 2 on right-hand axis), dual-scale visualization
officecli add "$FILE" "/6-Effects & Advanced" --type chart \
  --prop chartType=line \
  --prop title="Revenue vs Growth Rate" \
  --prop series1="Revenue:120,180,250,310,380,420" \
  --prop series2="Growth %:50,33,39,24,23,11" \
  --prop categories=2020,2021,2022,2023,2024,2025 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop secondaryAxis=2 \
  --prop colors=2E75B6,C00000 \
  --prop lineWidth=2.5 \
  --prop showMarkers=true --prop marker=circle:6:2E75B6 \
  --prop catTitle=Year --prop axisTitle=Revenue \
  --prop dataLabels=true --prop labelPos=top

# Chart 2: Reference line (target/threshold) with error bars
# referenceLine format: value:color:width:dash
#   - value: the threshold/target value on the Y axis
#   - color: hex RGB (no #)
#   - width: line thickness in pt (default 1.5)
#   - dash: solid/dot/dash/dashdot/longdash
# Features: referenceLine (horizontal target line), lineDash=longdash
officecli add "$FILE" "/6-Effects & Advanced" --type chart \
  --prop chartType=line \
  --prop title="vs Target (150)" \
  --prop dataRange=Sheet1!A1:C13 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=4472C4,70AD47 \
  --prop referenceLine=150:FF0000:1.5:dash \
  --prop showMarkers=true --prop marker=circle:4:4472C4 \
  --prop legend=bottom \
  --prop lineDash=longdash --prop lineWidth=1.5

# Chart 3: Title glow/shadow effects with per-series gradients
# Features: title.glow (color-radius-opacity), title.shadow,
#   series.shadow on line charts, plotFill + chartFill
officecli add "$FILE" "/6-Effects & Advanced" --type chart \
  --prop chartType=line \
  --prop title="Glow & Shadow Effects" \
  --prop series1="East:120,135,148,162,155,178" \
  --prop series2="West:110,118,130,145,138,162" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop lineWidth=3 --prop smooth=true \
  --prop colors=4472C4,ED7D31 \
  --prop title.glow=4472C4-8-60 \
  --prop title.shadow=000000-3-315-2-40 \
  --prop title.font=Calibri --prop title.size=16 \
  --prop title.bold=true --prop title.color=1F4E79 \
  --prop series.shadow=000000-3-315-1-30 \
  --prop plotFill=F0F4F8 --prop chartFill=FFFFFF

# Chart 4: Conditional coloring with chart/plot borders
# colorRule format: threshold:belowColor:aboveColor
#   - values below 0 → red (C00000), above 0 → green (70AD47)
# Features: colorRule (threshold-based conditional coloring),
#   chartArea.border, plotArea.border, referenceLine=0 (zero line)
officecli add "$FILE" "/6-Effects & Advanced" --type chart \
  --prop chartType=line \
  --prop title="Conditional Colors & Borders" \
  --prop series1="Profit:80,120,-30,160,-50,200,140,-20,180,90" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=2E75B6 \
  --prop lineWidth=2 \
  --prop showMarkers=true --prop marker=circle:6:2E75B6 \
  --prop colorRule=0:C00000:70AD47 \
  --prop referenceLine=0:888888:1:solid \
  --prop chartArea.border=D0D0D0:1:solid \
  --prop plotArea.border=E0E0E0:0.5:dot \
  --prop dataLabels=true --prop labelPos=top \
  --prop labelFont=8:666666:false

# ==========================================================================
# Sheet: 7-Line Elements
# ==========================================================================
echo "--- 7-Line Elements ---"
officecli add "$FILE" / --type sheet --prop name="7-Line Elements"

# Chart 1: Drop lines — vertical lines from data points to category axis
# Features: dropLines=true (simple toggle — default thin gray lines)
officecli add "$FILE" "/7-Line Elements" --type chart \
  --prop chartType=line \
  --prop title="Drop Lines" \
  --prop dataRange=Sheet1!A1:C13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=4472C4,ED7D31 \
  --prop showMarkers=true --prop marker=circle:5:4472C4 \
  --prop dropLines=true \
  --prop legend=bottom

# Chart 2: High-low lines — connect highest and lowest series at each point
# Features: hiLowLines=true (lines connecting highest and lowest values)
officecli add "$FILE" "/7-Line Elements" --type chart \
  --prop chartType=line \
  --prop title="High-Low Lines" \
  --prop series1="High:210,195,220,240,230,250" \
  --prop series2="Low:150,135,160,170,155,180" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=2E75B6,C00000 \
  --prop showMarkers=true --prop marker=diamond:5:2E75B6 \
  --prop hiLowLines=true \
  --prop legend=bottom

# Chart 3: Up-down bars with custom colors — show gain/loss between series
# updownbars format: gapWidth:upColor:downColor
#   - gapWidth: gap between bars (0-500, default 150)
#   - upColor: fill color for increase (Close > Open)
#   - downColor: fill color for decrease (Close < Open)
# Features: updownbars with custom colors (gain=green, loss=red)
officecli add "$FILE" "/7-Line Elements" --type chart \
  --prop chartType=line \
  --prop title="Up-Down Bars (Gain/Loss)" \
  --prop series1="Open:120,135,148,130,155,162" \
  --prop series2="Close:135,128,162,145,170,155" \
  --prop categories=Mon,Tue,Wed,Thu,Fri,Sat \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=4472C4,ED7D31 \
  --prop showMarkers=true --prop marker=circle:4:4472C4 \
  --prop updownbars=100:70AD47:C00000 \
  --prop legend=bottom

# Chart 4: Auto markers + 3D line with gapDepth
# Features: gapDepth=300 (3D depth spacing, 0-500), line3d with custom perspective
officecli add "$FILE" "/7-Line Elements" --type chart \
  --prop chartType=line3d \
  --prop title="3D Line with Gap Depth" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=15,25,30 \
  --prop gapDepth=300 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop chartFill=F5F5F5

# ==========================================================================
# Sheet: 8-Axis Extras
# ==========================================================================
echo "--- 8-Axis Extras ---"
officecli add "$FILE" / --type sheet --prop name="8-Axis Extras"

# Chart 1: crossesAt — value axis crosses category axis at specific value
# Features: crossesAt=50 (value axis crosses the category axis at y=50,
#   so bars/lines below 50 appear below the midline — great for threshold viz)
officecli add "$FILE" "/8-Axis Extras" --type chart \
  --prop chartType=line \
  --prop title="crossesAt — axis baseline at 50" \
  --prop series1=Score:40,65,55,80,45,90,70 \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun,Jul \
  --prop colors=2E75B6 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop crossesAt=50 \
  --prop lineWidth=2 --prop marker=circle --prop markerSize=6

# Chart 2: dispBlanksAs — how missing/null data points are rendered
# Features: dispBlanksAs=span (connect across null/blank data points with a
#   straight line — see also: gap=leave hole, zero=plot blank as 0)
officecli add "$FILE" "/8-Axis Extras" --type chart \
  --prop chartType=line \
  --prop title="dispBlanksAs=span (connect gaps)" \
  --prop series1=Revenue:100,120,130,150,160 \
  --prop categories=Jan,Feb,Mar,Apr,May \
  --prop colors=548235 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop dispBlanksAs=span \
  --prop lineWidth=2 --prop marker=circle --prop markerSize=6

# Chart 3: dispBlanksAs=zero + crossesAt=0
# Features: dispBlanksAs=zero (missing cells rendered as zero),
#   crossesAt=0 (axis crosses at y=0 — default for most charts).
#   Note: dispBlanksAs affects rendering when the data source has blank cells;
#   here it is shown as a metadata property on the chart (also accepts: gap, span).
officecli add "$FILE" "/8-Axis Extras" --type chart \
  --prop chartType=line \
  --prop title="dispBlanksAs=zero + crossesAt=0" \
  --prop series1=Revenue:100,120,130,150,160 \
  --prop categories=Jan,Feb,Mar,Apr,May \
  --prop colors=C00000 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop dispBlanksAs=zero \
  --prop crossesAt=0 \
  --prop lineWidth=2 --prop marker=circle --prop markerSize=6

# Chart 4: crosses=max (value axis on right side of plot)
# Features: crosses=max (value axis appears at the far end of the category
#   axis — i.e. on the right side for a left-to-right chart; also: autoZero,
#   min for the left/bottom edge)
officecli add "$FILE" "/8-Axis Extras" --type chart \
  --prop chartType=line \
  --prop title="crosses=max (value axis at far end)" \
  --prop series1=Index:45,60,52,75,80,68,90 \
  --prop categories=Mon,Tue,Wed,Thu,Fri,Sat,Sun \
  --prop colors=7030A0 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop crosses=max \
  --prop lineWidth=2

officecli close "$FILE"

officecli validate "$FILE"
echo "Generated: $FILE"
