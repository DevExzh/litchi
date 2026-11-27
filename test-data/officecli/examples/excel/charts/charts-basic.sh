#!/bin/bash
# Basic Charts Showcase — column, bar, line, and area charts with all variations.
# Generates: charts-basic.xlsx
#
# CLI twin of charts-basic.py (officecli Python SDK). Both produce an
# equivalent charts-basic.xlsx. See charts-basic.md for a guide to each sheet.
#
# Usage: ./charts-basic.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-basic.xlsx"
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

officecli set "$FILE" /Sheet1/A2 --prop text=Jan ; officecli set "$FILE" /Sheet1/B2 --prop text=120 ; officecli set "$FILE" /Sheet1/C2 --prop text=95  ; officecli set "$FILE" /Sheet1/D2 --prop text=88  ; officecli set "$FILE" /Sheet1/E2 --prop text=110
officecli set "$FILE" /Sheet1/A3 --prop text=Feb ; officecli set "$FILE" /Sheet1/B3 --prop text=135 ; officecli set "$FILE" /Sheet1/C3 --prop text=108 ; officecli set "$FILE" /Sheet1/D3 --prop text=92  ; officecli set "$FILE" /Sheet1/E3 --prop text=118
officecli set "$FILE" /Sheet1/A4 --prop text=Mar ; officecli set "$FILE" /Sheet1/B4 --prop text=148 ; officecli set "$FILE" /Sheet1/C4 --prop text=115 ; officecli set "$FILE" /Sheet1/D4 --prop text=105 ; officecli set "$FILE" /Sheet1/E4 --prop text=130
officecli set "$FILE" /Sheet1/A5 --prop text=Apr ; officecli set "$FILE" /Sheet1/B5 --prop text=162 ; officecli set "$FILE" /Sheet1/C5 --prop text=128 ; officecli set "$FILE" /Sheet1/D5 --prop text=118 ; officecli set "$FILE" /Sheet1/E5 --prop text=145
officecli set "$FILE" /Sheet1/A6 --prop text=May ; officecli set "$FILE" /Sheet1/B6 --prop text=155 ; officecli set "$FILE" /Sheet1/C6 --prop text=142 ; officecli set "$FILE" /Sheet1/D6 --prop text=125 ; officecli set "$FILE" /Sheet1/E6 --prop text=138
officecli set "$FILE" /Sheet1/A7 --prop text=Jun ; officecli set "$FILE" /Sheet1/B7 --prop text=178 ; officecli set "$FILE" /Sheet1/C7 --prop text=155 ; officecli set "$FILE" /Sheet1/D7 --prop text=138 ; officecli set "$FILE" /Sheet1/E7 --prop text=162
officecli set "$FILE" /Sheet1/A8 --prop text=Jul ; officecli set "$FILE" /Sheet1/B8 --prop text=195 ; officecli set "$FILE" /Sheet1/C8 --prop text=168 ; officecli set "$FILE" /Sheet1/D8 --prop text=145 ; officecli set "$FILE" /Sheet1/E8 --prop text=175
officecli set "$FILE" /Sheet1/A9 --prop text=Aug ; officecli set "$FILE" /Sheet1/B9 --prop text=210 ; officecli set "$FILE" /Sheet1/C9 --prop text=175 ; officecli set "$FILE" /Sheet1/D9 --prop text=152 ; officecli set "$FILE" /Sheet1/E9 --prop text=190
officecli set "$FILE" /Sheet1/A10 --prop text=Sep ; officecli set "$FILE" /Sheet1/B10 --prop text=188 ; officecli set "$FILE" /Sheet1/C10 --prop text=160 ; officecli set "$FILE" /Sheet1/D10 --prop text=140 ; officecli set "$FILE" /Sheet1/E10 --prop text=170
officecli set "$FILE" /Sheet1/A11 --prop text=Oct ; officecli set "$FILE" /Sheet1/B11 --prop text=172 ; officecli set "$FILE" /Sheet1/C11 --prop text=148 ; officecli set "$FILE" /Sheet1/D11 --prop text=130 ; officecli set "$FILE" /Sheet1/E11 --prop text=155
officecli set "$FILE" /Sheet1/A12 --prop text=Nov ; officecli set "$FILE" /Sheet1/B12 --prop text=165 ; officecli set "$FILE" /Sheet1/C12 --prop text=135 ; officecli set "$FILE" /Sheet1/D12 --prop text=122 ; officecli set "$FILE" /Sheet1/E12 --prop text=148
officecli set "$FILE" /Sheet1/A13 --prop text=Dec ; officecli set "$FILE" /Sheet1/B13 --prop text=198 ; officecli set "$FILE" /Sheet1/C13 --prop text=158 ; officecli set "$FILE" /Sheet1/D13 --prop text=142 ; officecli set "$FILE" /Sheet1/E13 --prop text=180

# ==========================================================================
# Sheet: 1-Column Charts
# ==========================================================================
echo "--- 1-Column Charts ---"
officecli add "$FILE" / --type sheet --prop name="1-Column Charts"

# Chart 1: Basic clustered column from cell range with axis titles
# Features: chartType=column, dataRange, catTitle, axisTitle, axisfont, gridlines
officecli add "$FILE" "/1-Column Charts" --type chart \
  --prop chartType=column \
  --prop title="Regional Sales by Month" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop catTitle=Month --prop axisTitle=Sales \
  --prop axisfont=9:58626E:Arial \
  --prop gridlines=D9D9D9:0.5:dot

# Chart 2: Stacked column with custom colors, data labels, and gap control
# Features: columnStacked, colors, dataLabels, labelPos, gapwidth, series.outline
officecli add "$FILE" "/1-Column Charts" --type chart \
  --prop chartType=columnStacked \
  --prop title="Stacked Regional Sales" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop colors=2E75B6,70AD47,FFC000,C00000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=center \
  --prop gapwidth=60 \
  --prop series.outline=FFFFFF-0.5

# Chart 3: 100% stacked column with legend position and plotFill
# Features: columnPercentStacked, legend=bottom, legendfont, plotFill
officecli add "$FILE" "/1-Column Charts" --type chart \
  --prop chartType=columnPercentStacked \
  --prop title="Market Share by Month" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom \
  --prop legendfont=9:8B949E \
  --prop plotFill=F5F5F5

# Chart 4: 3D column with perspective and title styling
# Features: column3d, view3d (rotX,rotY,perspective), title.font/size/color/bold
officecli add "$FILE" "/1-Column Charts" --type chart \
  --prop chartType=column3d \
  --prop title="3D Regional Sales" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=15,20,30 \
  --prop title.font=Calibri --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true

# ==========================================================================
# Sheet: 2-Bar Charts
# ==========================================================================
echo "--- 2-Bar Charts ---"
officecli add "$FILE" / --type sheet --prop name="2-Bar Charts"

# Chart 1: Horizontal bar with inline data and gapwidth
# Features: bar, inline data (Name:v1;Name2:v2), gapwidth, labelPos=outsideEnd
officecli add "$FILE" "/2-Bar Charts" --type chart \
  --prop chartType=bar \
  --prop title="Q4 Sales by Region" \
  --prop 'data=East:198;South:158;North:142;West:180' \
  --prop categories=East,South,North,West \
  --prop colors=2E75B6,70AD47,FFC000,C00000 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop gapwidth=80 \
  --prop dataLabels=true --prop labelPos=outsideEnd

# Chart 2: Stacked bar with named series and overlap
# Features: barStacked, named series (series1=Name:v1,v2), overlap
officecli add "$FILE" "/2-Bar Charts" --type chart \
  --prop chartType=barStacked \
  --prop title="H1 vs H2 Sales" \
  --prop series1=H1:663,598,528,661 \
  --prop series2=H2:833,718,669,868 \
  --prop categories=East,South,North,West \
  --prop colors=4472C4,ED7D31 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=center \
  --prop gapwidth=50 --prop overlap=0

# Chart 3: 100% stacked bar with reference line
# Note: on a barPercentStacked chart, the value axis is 0-1 (displayed as 0%-100%),
# so a 50% reference line must be written as 0.5 — not 50.
# referenceLine supports: value | value:color | value:color:label
# | value:color:width:dash | value:color:label:dash (legacy)
# | value:color:width:dash:label (canonical). Width is in points; default 1.5pt.
# Features: barPercentStacked, referenceLine, axisLine, catAxisLine
officecli add "$FILE" "/2-Bar Charts" --type chart \
  --prop chartType=barPercentStacked \
  --prop title="Regional Contribution %" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop referenceLine=0.5:FF0000:Target:dash \
  --prop axisLine=333333:1:solid \
  --prop catAxisLine=333333:1:solid

# Chart 4: 3D bar with chart area fill and display units
# Features: bar3d, chartFill (chart area background), style/styleId (preset 1-48)
officecli add "$FILE" "/2-Bar Charts" --type chart \
  --prop chartType=bar3d \
  --prop title="3D Regional Comparison" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=10,30,20 \
  --prop chartFill=F2F2F2 \
  --prop style=3

# ==========================================================================
# Sheet: 3-Line Charts
# ==========================================================================
echo "--- 3-Line Charts ---"
officecli add "$FILE" / --type sheet --prop name="3-Line Charts"

# Chart 1: Line with markers and cell-range series (dotted syntax)
# Features: series.name/values/categories (cell range), marker (style:size:color),
#   gridlines, minorGridlines
officecli add "$FILE" "/3-Line Charts" --type chart \
  --prop chartType=line \
  --prop title="East Region Trend" \
  --prop series1.name=East \
  --prop series1.values=Sheet1!B2:B13 \
  --prop series1.categories=Sheet1!A2:A13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop showMarkers=true --prop marker=circle:6:2E75B6 \
  --prop gridlines=D9D9D9:0.5:dot \
  --prop minorGridlines=EEEEEE:0.3:dot

# Chart 2: Smooth line with custom width and no gridlines
# Features: smooth, lineWidth, gridlines=none, series.shadow (color-blur-angle-dist-opacity)
officecli add "$FILE" "/3-Line Charts" --type chart \
  --prop chartType=line \
  --prop title="Smoothed Sales Trend" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop smooth=true --prop lineWidth=2.5 \
  --prop colors=0070C0,00B050,FFC000,FF0000 \
  --prop gridlines=none \
  --prop series.shadow=000000-4-315-2-40

# Chart 3: Stacked line
# Features: lineStacked, majorTickMark, tickLabelPos
officecli add "$FILE" "/3-Line Charts" --type chart \
  --prop chartType=lineStacked \
  --prop title="Cumulative Sales" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop catTitle=Month --prop axisTitle=Cumulative \
  --prop majorTickMark=outside --prop tickLabelPos=low

# Chart 4: Line with dashed lines, data table, and hidden legend
# Features: lineDash (solid/dot/dash/dashdot/longdash), dataTable, legend=none
officecli add "$FILE" "/3-Line Charts" --type chart \
  --prop chartType=line \
  --prop title="Trend with Data Table" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop lineDash=dash --prop lineWidth=1.5 \
  --prop dataTable=true \
  --prop legend=none

# ==========================================================================
# Sheet: 4-Area Charts
# ==========================================================================
echo "--- 4-Area Charts ---"
officecli add "$FILE" / --type sheet --prop name="4-Area Charts"

# Chart 1: Area with transparency and gradient fill
# Features: area, transparency (0-100%), gradient (color1-color2:angle)
officecli add "$FILE" "/4-Area Charts" --type chart \
  --prop chartType=area \
  --prop title="Sales Volume" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop transparency=40 \
  --prop gradient=4472C4-BDD7EE:90

# Chart 2: Stacked area with plotFill and rounded corners
# Features: areaStacked, plotFill, roundedCorners
officecli add "$FILE" "/4-Area Charts" --type chart \
  --prop chartType=areaStacked \
  --prop title="Stacked Volume" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop plotFill=F5F5F5 \
  --prop roundedCorners=true \
  --prop transparency=30

# Chart 3: 100% stacked area with axis control
# Features: areaPercentStacked, axisVisible, axisLine
officecli add "$FILE" "/4-Area Charts" --type chart \
  --prop chartType=areaPercentStacked \
  --prop title="Regional Mix %" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop transparency=20 \
  --prop axisVisible=true \
  --prop axisLine=999999:0.5:solid

# Chart 4: 3D area with perspective
# Features: area3d, view3d
officecli add "$FILE" "/4-Area Charts" --type chart \
  --prop chartType=area3d \
  --prop title="3D Sales Volume" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=20,25,15 \
  --prop colors=5B9BD5,A5D5A5,FFD966,F4B183

# ==========================================================================
# Sheet: 5-Styling
# Demonstrates all styling/layout properties on a single column chart
# ==========================================================================
echo "--- 5-Styling ---"
officecli add "$FILE" / --type sheet --prop name="5-Styling"

# Chart 1: Fully styled column chart — title, legend, axis, series effects
# Features: title.font/size/color/bold/shadow, legendfont, axisfont,
#   series.outline, series.shadow, roundedCorners, referenceLine
officecli add "$FILE" "/5-Styling" --type chart \
  --prop chartType=column \
  --prop title="Fully Styled Chart" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=0 --prop y=0 --prop width=14 --prop height=20 \
  --prop title.font=Georgia --prop title.size=18 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop title.shadow=000000-3-315-2-30 \
  --prop legendfont=10:444444:Helvetica \
  --prop legend=right \
  --prop axisfont=9:58626E:Arial \
  --prop catTitle=Month --prop axisTitle=Revenue \
  --prop gridlines=CCCCCC:0.5:dot \
  --prop plotFill=FAFAFA \
  --prop chartFill=FFFFFF \
  --prop series.outline=FFFFFF-0.5 \
  --prop series.shadow=000000-3-315-2-25 \
  --prop gapwidth=100 \
  --prop roundedCorners=true \
  --prop referenceLine=160:FF0000:1:dash \
  --prop colors=4472C4,ED7D31,70AD47,FFC000

# Chart 2: Column with secondary axis (dual Y-axis)
# Features: secondaryAxis (comma-separated 1-based series indices for second Y-axis)
officecli add "$FILE" "/5-Styling" --type chart \
  --prop chartType=column \
  --prop title="Sales vs Growth Rate" \
  --prop series1=Sales:120,135,148,162 \
  --prop series2=Growth:5.2,8.1,12.3,15.6 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=15 --prop y=0 --prop width=10 --prop height=20 \
  --prop secondaryAxis=2 \
  --prop colors=4472C4,FF0000

# Chart 3: Column with individual point colors and inverted negatives
# Features: point{N}.color (per-point coloring), invertIfNeg
officecli add "$FILE" "/5-Styling" --type chart \
  --prop chartType=column \
  --prop title="Quarterly P&L" \
  --prop 'series1=P&L:500,300,-200,800' \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=0 --prop y=21 --prop width=10 --prop height=18 \
  --prop point1.color=70AD47 --prop point2.color=70AD47 \
  --prop point3.color=FF0000 --prop point4.color=70AD47 \
  --prop invertIfNeg=true \
  --prop dataLabels=true --prop labelPos=outsideEnd

# Chart 4: Line with gradient plot area and custom data labels
# Features: plotFill gradient (color1-color2:angle), marker styles (diamond),
#   dataLabels.numFmt, dataLabel{N}.text (custom text for one label)
officecli add "$FILE" "/5-Styling" --type chart \
  --prop chartType=line \
  --prop title="Custom Labels Demo" \
  --prop series1=Revenue:100,200,300,250 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=11 --prop y=21 --prop width=14 --prop height=18 \
  --prop plotFill=E8F0FE-FFFFFF:90 \
  --prop showMarkers=true --prop marker=diamond:8:4472C4 \
  --prop lineWidth=2 \
  --prop dataLabels=true --prop labelPos=top \
  --prop dataLabels.numFmt=#,##0 \
  --prop dataLabel3.text=Peak!

# ==========================================================================
# Sheet: 6-Layout
# Manual layout of plot area, title, legend; axis orientation; log scale;
# display units; label font and separator; error bars
# ==========================================================================
echo "--- 6-Layout ---"
officecli add "$FILE" / --type sheet --prop name="6-Layout"

# Chart 1: Manual layout positioning of plot area, title, legend
# Features: plotArea.x/y/w/h (0-1 fraction), title.x/y, legend.x/y, legend.overlay
officecli add "$FILE" "/6-Layout" --type chart \
  --prop chartType=column \
  --prop title="Manual Layout" \
  --prop dataRange=Sheet1!A1:C13 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop plotArea.x=0.15 --prop plotArea.y=0.15 \
  --prop plotArea.w=0.7 --prop plotArea.h=0.7 \
  --prop title.x=0.3 --prop title.y=0.01 \
  --prop legend.x=0.02 --prop legend.y=0.4 \
  --prop legend.overlay=true

# Chart 2: Reversed axis, log scale, display units
# Features: logBase (logarithmic scale), axisOrientation=maxMin (reversed),
#   dispUnits (thousands/millions)
officecli add "$FILE" "/6-Layout" --type chart \
  --prop chartType=bar \
  --prop title="Log Scale + Reversed Axis" \
  --prop series1=Revenue:10,100,1000,10000 \
  --prop categories=Startup,Small,Medium,Enterprise \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop logBase=10 \
  --prop axisOrientation=maxMin \
  --prop dispUnits=thousands

# Chart 3: Label font, separator, leader lines, and per-label layout
# Features: labelFont (size:color:bold), dataLabels.separator,
#   dataLabel{N}.text (custom), dataLabel{N}.delete (hide one label)
officecli add "$FILE" "/6-Layout" --type chart \
  --prop chartType=column \
  --prop title="Label Formatting" \
  --prop series1=Sales:120,200,150,180 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop labelFont=11:2E75B6:true \
  --prop 'dataLabels.separator=: ' \
  --prop dataLabel2.text=Best! \
  --prop dataLabel3.delete=true

# Chart 4: Error bars, minor ticks, opacity
# Features: errBars (percentage/stdDev/fixed), minorTickMark, opacity (0-100%)
officecli add "$FILE" "/6-Layout" --type chart \
  --prop chartType=line \
  --prop title="Error Bars + Ticks" \
  --prop series1=Measurement:50,55,48,62,58 \
  --prop categories=Mon,Tue,Wed,Thu,Fri \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop showMarkers=true --prop marker=square:7:4472C4 \
  --prop errBars=percentage \
  --prop majorTickMark=outside --prop minorTickMark=inside \
  --prop opacity=80

# ==========================================================================
# Sheet: 7-Effects
# Gradients, conditional color, area fill, title glow, preset themes
# ==========================================================================
echo "--- 7-Effects ---"
officecli add "$FILE" / --type sheet --prop name="7-Effects"

# Chart 1: Per-series gradients
# Features: gradients (per-series, semicolon-separated "C1-C2:angle")
officecli add "$FILE" "/7-Effects" --type chart \
  --prop chartType=column \
  --prop title="Per-Series Gradients" \
  --prop series1=East:120,135,148 \
  --prop series2=West:110,118,130 \
  --prop categories=Q1,Q2,Q3 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90'

# Chart 2: Area fill gradient and title glow effect
# Features: areafill (area gradient), title.glow (color-radius-opacity)
officecli add "$FILE" "/7-Effects" --type chart \
  --prop chartType=area \
  --prop title="Glow Title + Area Fill" \
  --prop dataRange=Sheet1!A1:C13 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop areafill=4472C4-BDD7EE:90 \
  --prop transparency=30 \
  --prop title.glow=4472C4-8-60 \
  --prop title.size=16

# Chart 3: Conditional coloring rule
# Features: colorRule (threshold:belowColor:aboveColor — values below 60 red, above green)
officecli add "$FILE" "/7-Effects" --type chart \
  --prop chartType=column \
  --prop title="Conditional Colors" \
  --prop series1=Score:85,42,91,38,76,55 \
  --prop categories=A,B,C,D,E,F \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop colorRule=60:FF0000:70AD47 \
  --prop dataLabels=true --prop labelPos=outsideEnd

# Chart 4: Preset style/theme and leader lines
# Features: style (preset 1-48), dataLabels.showLeaderLines
officecli add "$FILE" "/7-Effects" --type chart \
  --prop chartType=column \
  --prop title="Preset Style 26" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop style=26 \
  --prop dataLabels=true \
  --prop dataLabels.showLeaderLines=true

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
echo "  8 sheets (Sheet1 data + 7 chart sheets, 28 charts total)"
