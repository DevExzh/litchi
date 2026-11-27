#!/bin/bash
# Column & Bar Charts Showcase — column, columnStacked, columnPercentStacked, and column3d.
# CLI twin of charts-column.py (officecli Python SDK). Both produce an equivalent
# charts-column.xlsx.
#
# Every column chart feature officecli supports is demonstrated at least once:
# gap width, overlap, bar shapes, axis scaling, gridlines, data labels,
# legend positioning, reference lines, secondary axis, gradients,
# transparency, shadows, manual layout, and 3D rotation.
#
# 8 sheets (Sheet1 data + 7 chart sheets), 28 charts total.
#
# Usage:
#   ./charts-column.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-column.xlsx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Source data — shared across all charts (Sheet1)
# ==========================================================================
officecli set "$FILE" /Sheet1/A1 --prop text=Month --prop bold=true
officecli set "$FILE" /Sheet1/B1 --prop text=East  --prop bold=true
officecli set "$FILE" /Sheet1/C1 --prop text=South --prop bold=true
officecli set "$FILE" /Sheet1/D1 --prop text=North --prop bold=true
officecli set "$FILE" /Sheet1/E1 --prop text=West  --prop bold=true

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
# Sheet: 1-Column Fundamentals
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="1-Column Fundamentals"

# Chart 1: Basic column with dataRange and axis titles
# Features: chartType=column, dataRange, catTitle, axisTitle, axisfont, gridlines, colors
officecli add "$FILE" "/1-Column Fundamentals" --type chart \
  --prop chartType=column \
  --prop title="Monthly Sales by Region" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop catTitle=Month --prop axisTitle=Revenue \
  --prop axisfont=9:58626E:Arial \
  --prop gridlines=D9D9D9:0.5:dot \
  --prop colors=4472C4,ED7D31,70AD47,FFC000

# Chart 2: Inline series with custom colors and gap width
# Features: inline series (series1=Name:v1,v2,...), colors, gapwidth, legend=bottom
officecli add "$FILE" "/1-Column Fundamentals" --type chart \
  --prop chartType=column \
  --prop title="Q1 Product Sales" \
  --prop series1="Laptops:320,280,350,310" \
  --prop series2="Phones:450,420,480,460" \
  --prop series3="Tablets:180,160,200,190" \
  --prop categories=Jan,Feb,Mar,Apr \
  --prop colors=2E75B6,C00000,70AD47 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop gapwidth=80 \
  --prop legend=bottom

# Chart 3: Dotted syntax with cell ranges
# Features: series.name/values/categories (cell range via dotted syntax), minorGridlines
officecli add "$FILE" "/1-Column Fundamentals" --type chart \
  --prop chartType=column \
  --prop title="East vs South (Cell Range)" \
  --prop series1.name=East \
  --prop series1.values=Sheet1!B2:B13 \
  --prop series1.categories=Sheet1!A2:A13 \
  --prop series2.name=South \
  --prop series2.values=Sheet1!C2:C13 \
  --prop series2.categories=Sheet1!A2:A13 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=4472C4,ED7D31 \
  --prop gridlines=D9D9D9:0.5:dot \
  --prop minorGridlines=EEEEEE:0.3:dot

# Chart 4: data= shorthand format
# Features: data (inline shorthand Name:v1;Name2:v2), legend=right
officecli add "$FILE" "/1-Column Fundamentals" --type chart \
  --prop chartType=column \
  --prop title="Weekly Output" \
  --prop 'data=Team A:85,92,78,95,88;Team B:70,80,85,90,75' \
  --prop categories=Mon,Tue,Wed,Thu,Fri \
  --prop colors=0070C0,FF6600 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=right

# ==========================================================================
# Sheet: 2-Column Variants
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="2-Column Variants"

# Chart 1: Stacked column with center data labels and series outline
# Features: columnStacked, dataLabels=center, series.outline
officecli add "$FILE" "/2-Column Variants" --type chart \
  --prop chartType=columnStacked \
  --prop title="Stacked Sales by Region" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop dataLabels=center \
  --prop series.outline=FFFFFF-0.5 \
  --prop legend=bottom

# Chart 2: 100% stacked column with axis number format
# Features: columnPercentStacked, axisNumFmt=0%, legend=bottom
officecli add "$FILE" "/2-Column Variants" --type chart \
  --prop chartType=columnPercentStacked \
  --prop title="Regional Contribution %" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=1F4E79,2E75B6,9DC3E6,BDD7EE \
  --prop axisNumFmt=0% \
  --prop legend=bottom \
  --prop gridlines=E0E0E0:0.5:solid

# Chart 3: 3D column with perspective and style
# Features: column3d, view3d (rotX,rotY,perspective), style (preset 1-48)
officecli add "$FILE" "/2-Column Variants" --type chart \
  --prop chartType=column3d \
  --prop title="3D Regional Trends" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=15,20,30 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop chartFill=F8F8F8 \
  --prop style=3

# Chart 4: 3D stacked column with gap depth
# Features: column3d stacked, gapDepth=200 (3D depth spacing)
officecli add "$FILE" "/2-Column Variants" --type chart \
  --prop chartType=column3d \
  --prop title="3D Stacked with Gap Depth" \
  --prop series1="East:120,135,148,162,155,178" \
  --prop series2="South:95,108,115,128,142,155" \
  --prop series3="North:88,92,105,118,125,138" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=15,20,30 \
  --prop gapDepth=200 \
  --prop colors=2E75B6,ED7D31,70AD47 \
  --prop legend=right

# ==========================================================================
# Sheet: 3-Column Styling
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="3-Column Styling"

# Chart 1: Title styling — font, size, color, bold
# Features: title.font, title.size, title.color, title.bold
officecli add "$FILE" "/3-Column Styling" --type chart \
  --prop chartType=column \
  --prop title="Styled Title Demo" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop legend=bottom

# Chart 2: Series shadow and outline effects
# Features: series.shadow (color-blur-angle-dist-opacity), series.outline (color-width)
officecli add "$FILE" "/3-Column Styling" --type chart \
  --prop chartType=column \
  --prop title="Shadow & Outline Effects" \
  --prop series1="Revenue:320,280,350,310,340" \
  --prop series2="Cost:210,195,230,220,215" \
  --prop categories=Q1,Q2,Q3,Q4,Q5 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=4472C4,C00000 \
  --prop series.shadow=000000-4-315-2-40 \
  --prop series.outline=FFFFFF-0.5 \
  --prop gapwidth=100 \
  --prop legend=bottom

# Chart 3: Per-series gradient fills
# Features: gradients (per-series gradient fills, start-end:angle)
officecli add "$FILE" "/3-Column Styling" --type chart \
  --prop chartType=column \
  --prop title="Gradient Columns" \
  --prop series1="East:120,135,148,162" \
  --prop series2="South:95,108,115,128" \
  --prop series3="North:88,92,105,118" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90;70AD47-C5E0B4:90' \
  --prop legend=bottom

# Chart 4: Transparency + plotFill gradient + chartFill + roundedCorners
# Features: transparency=30, plotFill gradient, chartFill, roundedCorners
officecli add "$FILE" "/3-Column Styling" --type chart \
  --prop chartType=column \
  --prop title="Transparent Columns on Gradient" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop transparency=30 \
  --prop plotFill=F0F4F8-D6E4F0:90 \
  --prop chartFill=FFFFFF \
  --prop colors=1F4E79,2E75B6,5B9BD5,9DC3E6 \
  --prop roundedCorners=true \
  --prop legend=bottom

# ==========================================================================
# Sheet: 4-Axis & Gridlines
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="4-Axis & Gridlines"

# Chart 1: Custom axis scaling — min, max, majorUnit, minorUnit
# Features: axisMin, axisMax, majorUnit, minorUnit, axisLine, catAxisLine
officecli add "$FILE" "/4-Axis & Gridlines" --type chart \
  --prop chartType=column \
  --prop title="Custom Axis Scale (50-250)" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop axisMin=50 --prop axisMax=250 --prop majorUnit=50 \
  --prop minorUnit=25 \
  --prop gridlines=D0D0D0:0.5:solid \
  --prop minorGridlines=EEEEEE:0.3:dot \
  --prop axisLine=C00000:1.5:solid \
  --prop catAxisLine=2E75B6:1.5:solid \
  --prop colors=4472C4,ED7D31,70AD47,FFC000

# Chart 2: Logarithmic scale with reversed axis
# Features: logBase=10 (logarithmic scale), axisReverse=true
officecli add "$FILE" "/4-Axis & Gridlines" --type chart \
  --prop chartType=column \
  --prop title="Log Scale (Base 10)" \
  --prop series1="Growth:1,10,100,1000,5000" \
  --prop categories="Year 1,Year 2,Year 3,Year 4,Year 5" \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop logBase=10 \
  --prop axisReverse=true \
  --prop colors=C00000 \
  --prop axisTitle="Value (log)" \
  --prop catTitle=Year \
  --prop gridlines=E0E0E0:0.5:dash

# Chart 3: Display units and axis number format
# Features: dispUnits=thousands, axisNumFmt=#,##0, majorTickMark=outside, minorTickMark=inside
officecli add "$FILE" "/4-Axis & Gridlines" --type chart \
  --prop chartType=column \
  --prop title="Revenue (in Thousands)" \
  --prop series1="Revenue:12000,18500,22000,31000,45000,52000" \
  --prop series2="Cost:8000,11000,14000,19500,28000,33000" \
  --prop categories=2020,2021,2022,2023,2024,2025 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop dispUnits=thousands \
  --prop axisNumFmt=#,##0 \
  --prop colors=2E75B6,C00000 \
  --prop catTitle=Year --prop axisTitle="Amount (K)" \
  --prop majorTickMark=outside --prop minorTickMark=inside \
  --prop legend=bottom

# Chart 4: Hidden axes with data table
# Features: gridlines=none, axisVisible=false, dataTable=true, legend=none
officecli add "$FILE" "/4-Axis & Gridlines" --type chart \
  --prop chartType=column \
  --prop title="Minimal Chart with Data Table" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop gridlines=none \
  --prop axisVisible=false \
  --prop dataTable=true \
  --prop legend=none \
  --prop colors=4472C4,ED7D31,70AD47,FFC000

# ==========================================================================
# Sheet: 5-Labels & Legend
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="5-Labels & Legend"

# Chart 1: Data labels with number format and styled label font
# Features: dataLabels=true, labelPos=outsideEnd, labelFont (size:color:bold), dataLabels.numFmt
officecli add "$FILE" "/5-Labels & Legend" --type chart \
  --prop chartType=column \
  --prop title="Sales with Labels" \
  --prop series1="Revenue:120,180,210,250,280" \
  --prop categories=Jan,Feb,Mar,Apr,May \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=4472C4 \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop labelFont=9:333333:true \
  --prop dataLabels.numFmt=#,##0

# Chart 2: Custom individual labels — delete some, highlight peak
# Features: dataLabel{N}.delete, dataLabel{N}.text, point{N}.color
officecli add "$FILE" "/5-Labels & Legend" --type chart \
  --prop chartType=column \
  --prop title="Peak Highlight" \
  --prop series1="Sales:88,120,165,210,195,178" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=2E75B6 \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop dataLabel1.delete=true --prop dataLabel2.delete=true \
  --prop dataLabel3.delete=true \
  --prop point4.color=C00000 \
  --prop dataLabel4.text=Peak! \
  --prop dataLabel5.delete=true --prop dataLabel6.delete=true

# Chart 3: Legend positioning and overlay with styled legend font
# Features: legend=right, legend.overlay=true, legendfont (size:color:fontname), plotFill
officecli add "$FILE" "/5-Labels & Legend" --type chart \
  --prop chartType=column \
  --prop title="Legend Overlay on Chart" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop legend=right \
  --prop legend.overlay=true \
  --prop legendfont=10:333333:Calibri \
  --prop plotFill=F5F5F5

# Chart 4: Manual layout — plotArea, title, and legend positioning
# Features: plotArea.x/y/w/h, title.x/y, legend.x/y/w/h (manual layout)
officecli add "$FILE" "/5-Labels & Legend" --type chart \
  --prop chartType=column \
  --prop title="Manual Layout Control" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
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
officecli add "$FILE" / --type sheet --prop name="6-Effects & Advanced"

# Chart 1: Secondary axis — dual Y-axis
# Features: secondaryAxis=2 (series 2 on right-hand axis)
officecli add "$FILE" "/6-Effects & Advanced" --type chart \
  --prop chartType=column \
  --prop title="Revenue vs Growth Rate" \
  --prop series1="Revenue:120,180,250,310,380,420" \
  --prop series2="Growth %:50,33,39,24,23,11" \
  --prop categories=2020,2021,2022,2023,2024,2025 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop secondaryAxis=2 \
  --prop colors=2E75B6,C00000 \
  --prop catTitle=Year --prop axisTitle=Revenue \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop legend=bottom

# Chart 2: Reference line (target/threshold)
# referenceLine format: value:color:width:dash
# Features: referenceLine (horizontal target line)
officecli add "$FILE" "/6-Effects & Advanced" --type chart \
  --prop chartType=column \
  --prop title="vs Target (150)" \
  --prop dataRange=Sheet1!A1:C13 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=4472C4,70AD47 \
  --prop referenceLine=150:FF0000:1.5:dash \
  --prop legend=bottom

# Chart 3: Title glow and shadow effects
# Features: title.glow (color-radius-opacity), title.shadow, series.shadow on column charts
officecli add "$FILE" "/6-Effects & Advanced" --type chart \
  --prop chartType=column \
  --prop title="Glow & Shadow Effects" \
  --prop series1="East:120,135,148,162,155,178" \
  --prop series2="West:110,118,130,145,138,162" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=4472C4,ED7D31 \
  --prop title.glow=4472C4-8-60 \
  --prop title.shadow=000000-3-315-2-40 \
  --prop title.font=Calibri --prop title.size=16 \
  --prop title.bold=true --prop title.color=1F4E79 \
  --prop series.shadow=000000-3-315-1-30 \
  --prop plotFill=F0F4F8 --prop chartFill=FFFFFF

# Chart 4: Conditional coloring with chart/plot borders
# colorRule format: threshold:belowColor:aboveColor
# Features: colorRule, chartArea.border, plotArea.border
officecli add "$FILE" "/6-Effects & Advanced" --type chart \
  --prop chartType=column \
  --prop title="Profit: Conditional Colors" \
  --prop series1="Profit:80,120,-30,160,-50,200,140,-20,180,90" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=2E75B6 \
  --prop colorRule=0:C00000:70AD47 \
  --prop referenceLine=0:888888:1:solid \
  --prop chartArea.border=D0D0D0:1:solid \
  --prop plotArea.border=E0E0E0:0.5:dot \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop labelFont=8:666666:false

# ==========================================================================
# Sheet: 7-Bar Shape & Gap
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="7-Bar Shape & Gap"

# Chart 1: Narrow gap width (bars close together)
# Features: gapwidth=30 (narrow gaps between column groups)
officecli add "$FILE" "/7-Bar Shape & Gap" --type chart \
  --prop chartType=column \
  --prop title="Narrow Gap (30%)" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop gapwidth=30 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop legend=bottom

# Chart 2: Wide gap with negative overlap (separated bars within group)
# Features: gapwidth=200 (wide gap), overlap=-50 (negative = bars separated)
officecli add "$FILE" "/7-Bar Shape & Gap" --type chart \
  --prop chartType=column \
  --prop title="Wide Gap + Negative Overlap" \
  --prop dataRange=Sheet1!A1:E7 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop gapwidth=200 \
  --prop overlap=-50 \
  --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
  --prop legend=bottom

# Chart 3: 3D column with cylinder shape
# Features: shape=cylinder (3D column bar shape)
officecli add "$FILE" "/7-Bar Shape & Gap" --type chart \
  --prop chartType=column3d \
  --prop title="Cylinder Shape" \
  --prop series1="East:120,135,148,162,155,178" \
  --prop series2="South:95,108,115,128,142,155" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop shape=cylinder \
  --prop view3d=15,20,30 \
  --prop colors=4472C4,ED7D31 \
  --prop legend=bottom

# Chart 4: 3D column with cone/pyramid shapes
# Features: shape=cone (3D column bar shape — also supports pyramid)
officecli add "$FILE" "/7-Bar Shape & Gap" --type chart \
  --prop chartType=column3d \
  --prop title="Cone Shape" \
  --prop series1="North:88,92,105,118,125,138" \
  --prop series2="West:110,118,130,145,138,162" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop shape=cone \
  --prop view3d=15,20,30 \
  --prop colors=70AD47,FFC000 \
  --prop legend=bottom

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
