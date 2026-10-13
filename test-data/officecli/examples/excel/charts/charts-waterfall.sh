#!/bin/bash
# Waterfall Charts Showcase — generates charts-waterfall.xlsx exercising the
# xlsx waterfall chart type with all variations.
#
# CLI twin of charts-waterfall.py (officecli Python SDK). Both produce an
# equivalent charts-waterfall.xlsx.
#
# 4 sheets, 16 charts total.
#
# Usage: ./charts-waterfall.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-waterfall.xlsx"
rm -f "$FILE"

# Forward-compat: a few props aren't consumed by the chart handler yet —
# officecli warns and exits 2 (unsupported_property) but still creates the
# element. Tolerate that exit code so the showcase runs end-to-end, matching the
# SDK twin (whose batch doesn't abort on it). Any OTHER non-zero exit is a real
# error and aborts.
officecli() {
  command officecli "$@" || { rc=$?; [ "$rc" -eq 2 ] && return 0; return $rc; }
}

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Sheet: 1-Waterfall Fundamentals
# ==========================================================================
echo "--- 1-Waterfall Fundamentals ---"
officecli add "$FILE" / --type sheet --prop name="1-Waterfall Fundamentals"

# Chart 1: Basic P&L waterfall with increase/decrease/total colors
# Features: chartType=waterfall, data= name:value pairs, increaseColor,
#   decreaseColor, totalColor, dataLabels
officecli add "$FILE" "/1-Waterfall Fundamentals" --type chart \
  --prop chartType=waterfall \
  --prop title="P&L Summary" \
  --prop data=Start:1000,Revenue:500,Costs:-300,Tax:-100,Net:1100 \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true

# Chart 2: Budget waterfall with blue/red/amber theme and legend
# Features: waterfall legend=bottom, alternative color palette (blue/red/amber)
officecli add "$FILE" "/1-Waterfall Fundamentals" --type chart \
  --prop chartType=waterfall \
  --prop title="Budget vs Actual" \
  --prop data=Budget:5000,Sales:2000,Marketing:-800,Ops:-600,Net:5600 \
  --prop increaseColor=2E75B6 \
  --prop decreaseColor=C00000 \
  --prop totalColor=FFC000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop legend=bottom

# Chart 3: Quarterly cash flow bridge with more data points
# Features: waterfall with 10 categories (extended data points),
#   quarterly granularity
officecli add "$FILE" "/1-Waterfall Fundamentals" --type chart \
  --prop chartType=waterfall \
  --prop title="Quarterly Cash Flow" \
  --prop "data=Opening:3000,Q1 Sales:1200,Q1 Costs:-500,Q2 Sales:1500,Q2 Costs:-700,Q3 Sales:800,Q3 Costs:-400,Q4 Sales:2000,Q4 Costs:-900,Closing:6000" \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=ED7D31 \
  --prop totalColor=4472C4 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop dataLabels=true

# Chart 4: Waterfall with custom title styling
# Features: title.font, title.size, title.color, title.bold
officecli add "$FILE" "/1-Waterfall Fundamentals" --type chart \
  --prop chartType=waterfall \
  --prop title="Revenue Bridge" \
  --prop "data=Base:2500,New Clients:800,Upsell:400,Churn:-600,Total:3100" \
  --prop increaseColor=548235 \
  --prop decreaseColor=BF0000 \
  --prop totalColor=2F5496 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true

# ==========================================================================
# Sheet: 2-Waterfall Styling
# ==========================================================================
echo "--- 2-Waterfall Styling ---"
officecli add "$FILE" / --type sheet --prop name="2-Waterfall Styling"

# Chart 1: Title styling with font, size, color, bold, and shadow
# Features: title.font, title.size, title.color, title.bold, title.shadow
officecli add "$FILE" "/2-Waterfall Styling" --type chart \
  --prop chartType=waterfall \
  --prop title="Styled Title Demo" \
  --prop data=Start:800,Income:300,Expenses:-200,Net:900 \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop "title.font=Trebuchet MS" --prop title.size=18 \
  --prop title.color=833C0B --prop title.bold=true \
  --prop title.shadow=000000-3-315-2-30

# Chart 2: Series shadow, plotFill, chartFill, roundedCorners
# Features: series.shadow, plotFill, chartFill, roundedCorners
officecli add "$FILE" "/2-Waterfall Styling" --type chart \
  --prop chartType=waterfall \
  --prop title="Shadow & Fill Effects" \
  --prop data=Baseline:1500,Growth:600,Decline:-400,Result:1700 \
  --prop increaseColor=2E75B6 \
  --prop decreaseColor=C00000 \
  --prop totalColor=FFC000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop series.shadow=000000-4-315-2-30 \
  --prop plotFill=F0F0F0 \
  --prop chartFill=FAFAFA \
  --prop roundedCorners=true

# Chart 3: Gridlines styling and axis font
# Features: gridlineColor, axisfont (size:color:font)
officecli add "$FILE" "/2-Waterfall Styling" --type chart \
  --prop chartType=waterfall \
  --prop title="Gridlines & Axis Font" \
  --prop data=Open:2000,Add:750,Remove:-350,Close:2400 \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop gridlineColor=CCCCCC \
  --prop axisfont=10:333333:Calibri

# Chart 4: Chart area border and plot area border
# Features: chartArea.border (color-width), plotArea.border
officecli add "$FILE" "/2-Waterfall Styling" --type chart \
  --prop chartType=waterfall \
  --prop title="Border Styling" \
  --prop data=Initial:1200,Gain:500,Loss:-300,Final:1400 \
  --prop increaseColor=548235 \
  --prop decreaseColor=BF0000 \
  --prop totalColor=2F5496 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop chartArea.border=4472C4:2 \
  --prop plotArea.border=A5A5A5:1

# ==========================================================================
# Sheet: 3-Waterfall Labels & Axis
# ==========================================================================
echo "--- 3-Waterfall Labels & Axis ---"
officecli add "$FILE" / --type sheet --prop name="3-Waterfall Labels & Axis"

# Chart 1: Data labels with labelFont and numFmt
# Features: dataLabels, labelFont (size:color:bold), dataLabels.numFmt
officecli add "$FILE" "/3-Waterfall Labels & Axis" --type chart \
  --prop chartType=waterfall \
  --prop title="Labels with NumFmt" \
  --prop data=Start:4500,Revenue:1800,COGS:-1200,SGA:-600,Net:4500 \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true \
  --prop labelFont=10:333333:true \
  --prop dataLabels.numFmt=#,##0

# Chart 2: Axis min/max and majorUnit
# Features: axisMin, axisMax, majorUnit
officecli add "$FILE" "/3-Waterfall Labels & Axis" --type chart \
  --prop chartType=waterfall \
  --prop title="Custom Axis Range" \
  --prop data=Base:2000,Up:800,Down:-500,Total:2300 \
  --prop increaseColor=2E75B6 \
  --prop decreaseColor=C00000 \
  --prop totalColor=FFC000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop axisMin=0 --prop axisMax=3500 --prop majorUnit=500

# Chart 3: Legend positioning and legendfont
# Features: legend=right, legendfont (size:color:font)
officecli add "$FILE" "/3-Waterfall Labels & Axis" --type chart \
  --prop chartType=waterfall \
  --prop title="Legend Styling" \
  --prop data=Begin:3000,Earned:1100,Spent:-700,End:3400 \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=right \
  --prop legendfont=10:1F4E79:Helvetica

# Chart 4: Manual layout with plotArea.x/y/w/h
# Features: plotArea.x/y/w/h (manual layout, fractional coordinates)
officecli add "$FILE" "/3-Waterfall Labels & Axis" --type chart \
  --prop chartType=waterfall \
  --prop title="Manual Plot Layout" \
  --prop data=Start:1800,Add:600,Sub:-400,End:2000 \
  --prop increaseColor=548235 \
  --prop decreaseColor=BF0000 \
  --prop totalColor=2F5496 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop plotArea.x=0.15 --prop plotArea.y=0.15 \
  --prop plotArea.w=0.75 --prop plotArea.h=0.70

# ==========================================================================
# Sheet: 4-Waterfall Advanced
# ==========================================================================
echo "--- 4-Waterfall Advanced ---"
officecli add "$FILE" / --type sheet --prop name="4-Waterfall Advanced"

# Chart 1: Waterfall with referenceLine
# Features: referenceLine (value:label-color-dash-width)
officecli add "$FILE" "/4-Waterfall Advanced" --type chart \
  --prop chartType=waterfall \
  --prop title="Reference Line" \
  --prop data=Start:2000,Revenue:900,Refunds:-300,Fees:-200,Net:2400 \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop referenceLine=2000:FF0000:Target:dash

# Chart 2: Axis line and category axis line styling
# Features: axisLine (color-width), catAxisLine
officecli add "$FILE" "/4-Waterfall Advanced" --type chart \
  --prop chartType=waterfall \
  --prop title="Axis Line Styling" \
  --prop data=Open:1500,Deposit:700,Withdraw:-400,Close:1800 \
  --prop increaseColor=2E75B6 \
  --prop decreaseColor=C00000 \
  --prop totalColor=FFC000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop axisLine=333333:2 \
  --prop catAxisLine=333333:2

# Chart 3: Title glow and shadow effects
# Features: title.glow (color-radius), title.shadow
officecli add "$FILE" "/4-Waterfall Advanced" --type chart \
  --prop chartType=waterfall \
  --prop title="Glow & Shadow Effects" \
  --prop data=Base:3000,Inflow:1200,Outflow:-800,Balance:3400 \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop title.glow=4472C4-8 \
  --prop title.shadow=000000-3-315-2-30 \
  --prop title.size=16 --prop title.bold=true

# Chart 4: Large dataset waterfall (8+ categories)
# Features: large dataset (12 categories), axisfont with smaller size for readability
officecli add "$FILE" "/4-Waterfall Advanced" --type chart \
  --prop chartType=waterfall \
  --prop title="Annual P&L Detail" \
  --prop "data=Revenue:8500,COGS:-3400,Gross Profit:5100,R&D:-1200,Sales:-900,Marketing:-600,G&A:-500,EBITDA:1900,Depreciation:-300,Interest:-200,Tax:-350,Net Income:1050" \
  --prop increaseColor=548235 \
  --prop decreaseColor=C00000 \
  --prop totalColor=2F5496 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop dataLabels=true \
  --prop axisfont=8:333333:Calibri

# Remove blank default Sheet1 (all data is inline)
officecli remove "$FILE" /Sheet1

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
