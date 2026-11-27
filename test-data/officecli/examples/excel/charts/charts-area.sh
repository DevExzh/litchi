#!/bin/bash
# Area Charts Showcase — area, areaStacked, areaPercentStacked, and area3d.
# CLI twin of charts-area.py (officecli Python SDK). Both produce an
# equivalent charts-area.xlsx.
#
# 5 sheets, 20 charts total. Every area chart feature officecli supports is
# demonstrated at least once: area fills, gradients, transparency, stacking,
# axis scaling, gridlines, data labels, legend positioning, reference lines,
# secondary axis, shadows, manual layout, and 3D rotation.
#
# Usage: ./charts-area.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-area.xlsx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Source data — shared across all charts
# ==========================================================================
officecli set "$FILE" /Sheet1/A1 --prop text=Month    --prop bold=true
officecli set "$FILE" /Sheet1/B1 --prop text=Organic  --prop bold=true
officecli set "$FILE" /Sheet1/C1 --prop text=Paid     --prop bold=true
officecli set "$FILE" /Sheet1/D1 --prop text=Social   --prop bold=true
officecli set "$FILE" /Sheet1/E1 --prop text=Referral --prop bold=true

officecli set "$FILE" /Sheet1/A2  --prop text=Jan
officecli set "$FILE" /Sheet1/B2  --prop text=4200
officecli set "$FILE" /Sheet1/C2  --prop text=3100
officecli set "$FILE" /Sheet1/D2  --prop text=1800
officecli set "$FILE" /Sheet1/E2  --prop text=1200
officecli set "$FILE" /Sheet1/A3  --prop text=Feb
officecli set "$FILE" /Sheet1/B3  --prop text=4800
officecli set "$FILE" /Sheet1/C3  --prop text=3500
officecli set "$FILE" /Sheet1/D3  --prop text=2100
officecli set "$FILE" /Sheet1/E3  --prop text=1400
officecli set "$FILE" /Sheet1/A4  --prop text=Mar
officecli set "$FILE" /Sheet1/B4  --prop text=5100
officecli set "$FILE" /Sheet1/C4  --prop text=3800
officecli set "$FILE" /Sheet1/D4  --prop text=2400
officecli set "$FILE" /Sheet1/E4  --prop text=1500
officecli set "$FILE" /Sheet1/A5  --prop text=Apr
officecli set "$FILE" /Sheet1/B5  --prop text=5600
officecli set "$FILE" /Sheet1/C5  --prop text=4200
officecli set "$FILE" /Sheet1/D5  --prop text=2800
officecli set "$FILE" /Sheet1/E5  --prop text=1700
officecli set "$FILE" /Sheet1/A6  --prop text=May
officecli set "$FILE" /Sheet1/B6  --prop text=6200
officecli set "$FILE" /Sheet1/C6  --prop text=4800
officecli set "$FILE" /Sheet1/D6  --prop text=3200
officecli set "$FILE" /Sheet1/E6  --prop text=1900
officecli set "$FILE" /Sheet1/A7  --prop text=Jun
officecli set "$FILE" /Sheet1/B7  --prop text=6800
officecli set "$FILE" /Sheet1/C7  --prop text=5200
officecli set "$FILE" /Sheet1/D7  --prop text=3600
officecli set "$FILE" /Sheet1/E7  --prop text=2100
officecli set "$FILE" /Sheet1/A8  --prop text=Jul
officecli set "$FILE" /Sheet1/B8  --prop text=7500
officecli set "$FILE" /Sheet1/C8  --prop text=5800
officecli set "$FILE" /Sheet1/D8  --prop text=4000
officecli set "$FILE" /Sheet1/E8  --prop text=2300
officecli set "$FILE" /Sheet1/A9  --prop text=Aug
officecli set "$FILE" /Sheet1/B9  --prop text=8100
officecli set "$FILE" /Sheet1/C9  --prop text=6300
officecli set "$FILE" /Sheet1/D9  --prop text=4300
officecli set "$FILE" /Sheet1/E9  --prop text=2500
officecli set "$FILE" /Sheet1/A10 --prop text=Sep
officecli set "$FILE" /Sheet1/B10 --prop text=7600
officecli set "$FILE" /Sheet1/C10 --prop text=5900
officecli set "$FILE" /Sheet1/D10 --prop text=3900
officecli set "$FILE" /Sheet1/E10 --prop text=2300
officecli set "$FILE" /Sheet1/A11 --prop text=Oct
officecli set "$FILE" /Sheet1/B11 --prop text=7200
officecli set "$FILE" /Sheet1/C11 --prop text=5500
officecli set "$FILE" /Sheet1/D11 --prop text=3500
officecli set "$FILE" /Sheet1/E11 --prop text=2100
officecli set "$FILE" /Sheet1/A12 --prop text=Nov
officecli set "$FILE" /Sheet1/B12 --prop text=6900
officecli set "$FILE" /Sheet1/C12 --prop text=5100
officecli set "$FILE" /Sheet1/D12 --prop text=3200
officecli set "$FILE" /Sheet1/E12 --prop text=1900
officecli set "$FILE" /Sheet1/A13 --prop text=Dec
officecli set "$FILE" /Sheet1/B13 --prop text=7800
officecli set "$FILE" /Sheet1/C13 --prop text=5700
officecli set "$FILE" /Sheet1/D13 --prop text=3800
officecli set "$FILE" /Sheet1/E13 --prop text=2200

# ==========================================================================
# Sheet: 1-Area Fundamentals
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="1-Area Fundamentals"

# Chart 1: Basic area chart with dataRange, axis titles, and custom colors
# Features: chartType=area, dataRange, colors, catTitle, axisTitle, gridlines
officecli add "$FILE" "/1-Area Fundamentals" --type chart \
  --prop chartType=area \
  --prop title="Website Traffic Overview" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop catTitle=Month --prop axisTitle=Visitors \
  --prop gridlines=D9D9D9:0.5:dot

# Chart 2: Inline series with transparency
# Features: inline series, transparency (0-100), legend=bottom
officecli add "$FILE" "/1-Area Fundamentals" --type chart \
  --prop chartType=area \
  --prop title="Quarterly Revenue Streams" \
  --prop series1="Subscriptions:120,180,210,250" \
  --prop series2="One-time:90,140,160,200" \
  --prop series3="Services:60,85,110,145" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop colors=2E75B6,70AD47,FFC000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop transparency=40 \
  --prop legend=bottom

# Chart 3: Area with areafill gradient
# Features: areafill (gradient from-to:angle), legend=none, single series
officecli add "$FILE" "/1-Area Fundamentals" --type chart \
  --prop chartType=area \
  --prop title="Monthly Active Users" \
  --prop series1="Users:3200,3800,4500,5100,5800,6400" \
  --prop categories=Jul,Aug,Sep,Oct,Nov,Dec \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop areafill=4472C4-BDD7EE:90 \
  --prop legend=none

# Chart 4: Per-series gradient fills
# Features: gradients (per-series gradient fills from-to:angle;...),
#   legendfont (size:color:font)
officecli add "$FILE" "/1-Area Fundamentals" --type chart \
  --prop chartType=area \
  --prop title="Revenue by Channel" \
  --prop series1="Direct:45,52,61,70" \
  --prop series2="Partner:30,38,42,55" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop "gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90" \
  --prop legend=right --prop legendfont=10:333333:Calibri

# ==========================================================================
# Sheet: 2-Area Variants
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="2-Area Variants"

# Chart 1: Stacked area with plotFill and rounded corners
# Features: chartType=areaStacked, plotFill (solid), roundedCorners
officecli add "$FILE" "/2-Area Variants" --type chart \
  --prop chartType=areaStacked \
  --prop title="Cumulative Traffic Sources" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop plotFill=F5F5F5 \
  --prop roundedCorners=true \
  --prop legend=bottom

# Chart 2: 100% stacked area with axis number format and axis line
# Features: chartType=areaPercentStacked, axisNumFmt, axisLine
officecli add "$FILE" "/2-Area Variants" --type chart \
  --prop chartType=areaPercentStacked \
  --prop title="Traffic Share by Channel" \
  --prop dataRange=Sheet1!A1:E13 \
  --prop colors=2E75B6,C55A11,548235,BF8F00 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop axisNumFmt=0% \
  --prop axisLine=333333:1:solid \
  --prop legend=bottom

# Chart 3: 3D area with perspective rotation
# Features: chartType=area3d, view3d (rotX,rotY,perspective)
officecli add "$FILE" "/2-Area Variants" --type chart \
  --prop chartType=area3d \
  --prop title="3D Regional Sales" \
  --prop series1="East:120,135,148,162,155,178" \
  --prop series2="West:95,108,115,128,142,155" \
  --prop series3="Central:88,92,105,118,125,138" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=20,25,15 \
  --prop legend=right

# Chart 4: 3D stacked area
# Features: area3d stacked appearance, multiple series, gridlines
officecli add "$FILE" "/2-Area Variants" --type chart \
  --prop chartType=area3d \
  --prop title="3D Stacked Inventory" \
  --prop series1="Warehouse A:500,480,520,550,530,560" \
  --prop series2="Warehouse B:320,350,340,380,400,410" \
  --prop series3="Warehouse C:180,200,210,230,250,240" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop colors=1F4E79,2E75B6,9DC3E6 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=15,20,20 \
  --prop gridlines=D9D9D9:0.5:dot

# ==========================================================================
# Sheet: 3-Area Styling
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="3-Area Styling"

# Chart 1: Title styling (font, size, color, bold, shadow)
# Features: title.font, title.size, title.color, title.bold, title.shadow
officecli add "$FILE" "/3-Area Styling" --type chart \
  --prop chartType=area \
  --prop title="Styled Title Demo" \
  --prop series1="Revenue:80,120,160,200,240,280" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop colors=4472C4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop title.shadow=000000-3-315-2-30 \
  --prop transparency=30

# Chart 2: Series shadow, outline, and smooth curve
# Features: smooth, series.shadow (color-blur-angle-dist-opacity),
#   series.outline (color-width)
officecli add "$FILE" "/3-Area Styling" --type chart \
  --prop chartType=area \
  --prop title="Smooth Area with Effects" \
  --prop series1="Signups:150,180,220,260,310,350" \
  --prop series2="Trials:90,110,140,170,200,230" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop colors=4472C4,70AD47 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop smooth=true \
  --prop series.shadow=000000-4-315-2-40 \
  --prop series.outline=333333-1 \
  --prop transparency=25

# Chart 3: Axis font styling, gridlines, and minor gridlines
# Features: axisfont (size:color:font), gridlines (color:width:dash),
#   minorGridlines
officecli add "$FILE" "/3-Area Styling" --type chart \
  --prop chartType=area \
  --prop title="Gridline Configuration" \
  --prop dataRange=Sheet1!A1:C13 \
  --prop colors=2E75B6,C55A11 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop axisfont=9:58626E:Arial \
  --prop gridlines=D9D9D9:0.5:dot \
  --prop minorGridlines=EEEEEE:0.3:dot \
  --prop catTitle=Month --prop axisTitle=Visitors

# Chart 4: Chart fill, plot fill gradient, chart/plot area borders
# Features: chartFill, plotFill (gradient from-to:angle),
#   chartArea.border, plotArea.border, roundedCorners
officecli add "$FILE" "/3-Area Styling" --type chart \
  --prop chartType=area \
  --prop title="Fills and Borders" \
  --prop series1="Sales:200,240,280,320,360,400" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop colors=4472C4 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop chartFill=FAFAFA \
  --prop "plotFill=E8F0FE-D6E4F0:90" \
  --prop chartArea.border=D0D0D0:1:solid \
  --prop plotArea.border=E0E0E0:0.5:dot \
  --prop roundedCorners=true

# ==========================================================================
# Sheet: 4-Labels & Legend
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="4-Labels & Legend"

# Chart 1: Data labels with position, font, and number format
# Features: dataLabels, labelPos (top), labelFont (size:color:bold),
#   dataLabels.numFmt
officecli add "$FILE" "/4-Labels & Legend" --type chart \
  --prop chartType=area \
  --prop title="Labeled Area Chart" \
  --prop series1="Users:3200,3800,4500,5100,5800,6400" \
  --prop categories=Jul,Aug,Sep,Oct,Nov,Dec \
  --prop colors=4472C4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=top \
  --prop labelFont=9:333333:true \
  --prop dataLabels.numFmt=#,##0

# Chart 2: Individual label deletion and per-point colors
# Features: dataLabel{N}.delete, point{N}.color
officecli add "$FILE" "/4-Labels & Legend" --type chart \
  --prop chartType=area \
  --prop title="Highlighted Peak Month" \
  --prop series1="Revenue:180,210,250,310,280,260" \
  --prop categories=Jul,Aug,Sep,Oct,Nov,Dec \
  --prop colors=2E75B6 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true \
  --prop dataLabel1.delete=true --prop dataLabel2.delete=true \
  --prop dataLabel5.delete=true --prop dataLabel6.delete=true \
  --prop point4.color=C00000 \
  --prop transparency=30

# Chart 3: Legend positioning with overlay and font styling
# Features: legend=right, legendfont, legend.overlay
officecli add "$FILE" "/4-Labels & Legend" --type chart \
  --prop chartType=area \
  --prop title="Legend Overlay Demo" \
  --prop series1="Desktop:4200,4800,5100,5600" \
  --prop series2="Mobile:3100,3500,3800,4200" \
  --prop series3="Tablet:1200,1400,1500,1700" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=right --prop legendfont=10:1F4E79:Calibri \
  --prop legend.overlay=true \
  --prop transparency=35

# Chart 4: Manual layout — plotArea positioning
# Features: plotArea.x/y/w/h, title.x/y, legend.x/y/w/h (manual layout)
officecli add "$FILE" "/4-Labels & Legend" --type chart \
  --prop chartType=area \
  --prop title="Manual Layout" \
  --prop series1="Growth:100,130,170,220,280,350" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop colors=70AD47 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop plotArea.x=0.12 --prop plotArea.y=0.18 \
  --prop plotArea.w=0.82 --prop plotArea.h=0.55 \
  --prop title.x=0.25 --prop title.y=0.02 \
  --prop legend.x=0.15 --prop legend.y=0.82 \
  --prop legend.w=0.7 --prop legend.h=0.12

# ==========================================================================
# Sheet: 5-Advanced
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="5-Advanced"

# Chart 1: Secondary axis (dual scale)
# Features: secondaryAxis (1-based series index on secondary Y axis)
officecli add "$FILE" "/5-Advanced" --type chart \
  --prop chartType=area \
  --prop title="Revenue vs Conversion Rate" \
  --prop series1="Revenue:120,180,250,310,280,340" \
  --prop series2="Conv %:2.1,2.8,3.2,3.9,3.5,4.1" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop colors=4472C4,C00000 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop secondaryAxis=2 \
  --prop transparency=30

# Chart 2: Reference line
# Features: referenceLine (value:color:width:dash)
officecli add "$FILE" "/5-Advanced" --type chart \
  --prop chartType=area \
  --prop title="Sales vs Target" \
  --prop series1="Sales:85,92,108,115,98,120" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop colors=4472C4 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop referenceLine=100:FF0000:1.5:dash \
  --prop transparency=25 \
  --prop areafill=4472C4-BDD7EE:90

# Chart 3: Axis min/max, major unit, log scale, display units
# Features: axisMin, axisMax, majorUnit, dispUnits (thousands/millions)
officecli add "$FILE" "/5-Advanced" --type chart \
  --prop chartType=area \
  --prop title="Axis Scaling Demo" \
  --prop series1="Visits:3200,3800,4500,5100,5800,6400" \
  --prop categories=Jul,Aug,Sep,Oct,Nov,Dec \
  --prop colors=2E75B6 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop axisMin=3000 --prop axisMax=7000 \
  --prop majorUnit=500 \
  --prop dispUnits=thousands \
  --prop "axisTitle=Visitors (K)" \
  --prop transparency=30

# Chart 4: Color rule, title glow, series shadow
# Features: colorRule (threshold:belowColor:aboveColor), title.glow
#   (color-radius-opacity), series.shadow
officecli add "$FILE" "/5-Advanced" --type chart \
  --prop chartType=area \
  --prop title="Performance Threshold" \
  --prop series1="Score:45,62,38,71,55,80" \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop colorRule=50:C00000:70AD47 \
  --prop referenceLine=50:888888:1:solid \
  --prop title.glow=4472C4-8-60 \
  --prop series.shadow=000000-3-315-1-30 \
  --prop transparency=20

officecli close "$FILE"

officecli validate "$FILE"
echo "Generated: $FILE"
