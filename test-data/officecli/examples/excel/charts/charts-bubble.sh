#!/bin/bash
# Bubble Charts Showcase — bubble scale, size representation, and styling.
# Generates charts-bubble.xlsx (4 chart sheets, 14 charts total).
#
# CLI twin of charts-bubble.py (officecli Python SDK). Both produce an
# equivalent charts-bubble.xlsx.
#
# Usage: ./charts-bubble.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-bubble.xlsx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Sheet: 1-Bubble Fundamentals
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="1-Bubble Fundamentals"

# Chart 1: Basic bubble chart with 2 series
# Features: chartType=bubble, X;Y;Size triplets, catTitle, axisTitle
officecli add "$FILE" "/1-Bubble Fundamentals" --type chart \
  --prop chartType=bubble \
  --prop title="Market Analysis" \
  --prop series1=Enterprise:80,45,60 \
  --prop series2=Consumer:50,35,70 \
  --prop colors=4472C4,ED7D31 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop "catTitle=Market Size" --prop "axisTitle=Growth Rate" \
  --prop legend=bottom

# Chart 2: bubbleScale=100 with dataLabels
# Features: bubbleScale=100, dataLabels with center positioning
officecli add "$FILE" "/1-Bubble Fundamentals" --type chart \
  --prop chartType=bubble \
  --prop title="Product Portfolio" \
  --prop series1=Products:90,50,70,40 \
  --prop colors=2E75B6 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop bubbleScale=100 \
  --prop dataLabels=true --prop labelPos=center \
  --prop labelFont=9:FFFFFF:true \
  --prop legend=bottom

# Chart 3: bubbleScale=50 (smaller bubbles)
# Features: bubbleScale=50
officecli add "$FILE" "/1-Bubble Fundamentals" --type chart \
  --prop chartType=bubble \
  --prop title="Small Bubbles (Scale 50)" \
  --prop series1=Tech:60,80,45 \
  --prop series2=Finance:55,70,35 \
  --prop colors=70AD47,FFC000 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop bubbleScale=50 \
  --prop legend=bottom

# Chart 4: sizeRepresents=width
# Features: sizeRepresents=width (bubble diameter proportional to value)
officecli add "$FILE" "/1-Bubble Fundamentals" --type chart \
  --prop chartType=bubble \
  --prop title="Size by Width" \
  --prop series1=Regions:70,40,55,85 \
  --prop colors=5B9BD5 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop sizeRepresents=width \
  --prop bubbleScale=100 \
  --prop legend=bottom

# ==========================================================================
# Sheet: 2-Bubble Styling
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="2-Bubble Styling"

# Chart 1: Title styling, legend positioning
# Features: title.font/size/color/bold, legend=right, legendfont
officecli add "$FILE" "/2-Bubble Styling" --type chart \
  --prop chartType=bubble \
  --prop title="Styled Bubble Chart" \
  --prop series1=SegmentA:65,50,80 \
  --prop series2=SegmentB:45,60,40 \
  --prop colors=1F4E79,C55A11 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop legend=right --prop legendfont=10:333333:Calibri

# Chart 2: Series colors, transparency
# Features: ARGB colors with alpha (80=50% transparency)
officecli add "$FILE" "/2-Bubble Styling" --type chart \
  --prop chartType=bubble \
  --prop title="Transparent Overlapping Bubbles" \
  --prop series1=GroupX:75,60,90,50 \
  --prop series2=GroupY:65,55,80,45 \
  --prop colors=804472C4,80ED7D31 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop bubbleScale=120 \
  --prop legend=bottom

# Chart 3: gridlines, axisfont, axisLine
# Features: gridlines, axisfont, axisLine
officecli add "$FILE" "/2-Bubble Styling" --type chart \
  --prop chartType=bubble \
  --prop title="Grid & Axis Styling" \
  --prop series1=Div1:55,70,45 \
  --prop series2=Div2:60,40,75 \
  --prop colors=2E75B6,548235 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop gridlines=D9D9D9:0.5 \
  --prop axisfont=9:666666 \
  --prop axisLine=333333:1 \
  --prop legend=bottom

# Chart 4: plotFill, chartFill, series.shadow
# Features: plotFill, chartFill, series.shadow
officecli add "$FILE" "/2-Bubble Styling" --type chart \
  --prop chartType=bubble \
  --prop title="Shadow & Fill Effects" \
  --prop series1=Portfolio:80,55,65,45 \
  --prop colors=4472C4 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop plotFill=F0F4F8 --prop chartFill=FAFAFA \
  --prop series.shadow=000000-4-315-2-30 \
  --prop legend=bottom

# ==========================================================================
# Sheet: 3-Bubble Advanced
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="3-Bubble Advanced"

# Chart 1: secondaryAxis
# Features: secondaryAxis on bubble chart
officecli add "$FILE" "/3-Bubble Advanced" --type chart \
  --prop chartType=bubble \
  --prop title="Dual-Axis Bubble" \
  --prop series1=Domestic:70,85,60,90 \
  --prop series2=International:45,55,80,65 \
  --prop categories=1,2,3,4 \
  --prop colors=4472C4,ED7D31 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop secondaryAxis=2 \
  --prop legend=bottom

# Chart 2: referenceLine
# Features: referenceLine on bubble chart
officecli add "$FILE" "/3-Bubble Advanced" --type chart \
  --prop chartType=bubble \
  --prop title="Growth Threshold" \
  --prop series1=Products:60,80,45,55 \
  --prop categories=1,2,3,4 \
  --prop colors=70AD47 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop "referenceLine=50:C00000:Target" \
  --prop bubbleScale=80 \
  --prop legend=bottom

# Chart 3: axisMin/Max, logBase
# Features: axisMin/Max, logBase=10 (logarithmic scale)
officecli add "$FILE" "/3-Bubble Advanced" --type chart \
  --prop chartType=bubble \
  --prop title="Log Scale Analysis" \
  --prop series1=Markets:5,15,50,120 \
  --prop categories=1,2,3,4 \
  --prop colors=2E75B6 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop axisMin=1 --prop axisMax=200 \
  --prop logBase=10 \
  --prop bubbleScale=80 \
  --prop legend=bottom

# Chart 4: chartArea.border, plotArea.border, trendline
# Features: chartArea.border, plotArea.border, trendline=linear
officecli add "$FILE" "/3-Bubble Advanced" --type chart \
  --prop chartType=bubble \
  --prop title="Trend & Borders" \
  --prop series1=Investments:20,55,95,140,180 \
  --prop categories=1,2,3,4,5 \
  --prop colors=4472C4 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop chartArea.border=333333:1.5 \
  --prop plotArea.border=999999:0.75 \
  --prop trendline=linear \
  --prop legend=bottom

# ==========================================================================
# Sheet: 4-Bubble Series Data
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="4-Bubble Series Data"

# Chart 1: shownegbubbles — render bubbles for negative size values
# Features: shownegbubbles=true (Excel hides negative-size bubbles by default)
officecli add "$FILE" "/4-Bubble Series Data" --type chart \
  --prop chartType=bubble \
  --prop title="shownegbubbles — negative sizes visible" \
  --prop series1=Data:60,30,90 \
  --prop series2=Neg:40,50,70 \
  --prop colors=4472C4,C00000 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop shownegbubbles=true \
  --prop bubbleScale=80 \
  --prop legend=bottom

# Chart 2: series1.bubbleSize (range ref) — sizes from worksheet cells
# Populate size data first, then reference it (bubbleSize + bubbleSizeRef round-trip).
officecli add "$FILE" "/4-Bubble Series Data" --type cell --prop ref=A1 --prop value=10
officecli add "$FILE" "/4-Bubble Series Data" --type cell --prop ref=A2 --prop value=25
officecli add "$FILE" "/4-Bubble Series Data" --type cell --prop ref=A3 --prop value=40
officecli add "$FILE" "/4-Bubble Series Data" --type chart \
  --prop chartType=bubble \
  --prop title="series1.bubbleSize — range ref" \
  --prop series1=Sizes:80,45,60 \
  --prop 'series1.bubbleSize=4-Bubble Series Data!$A$1:$A$3' \
  --prop colors=70AD47 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop bubbleScale=100 --prop legend=bottom

# Remove blank default Sheet1 (all data is inline)
officecli remove "$FILE" /Sheet1

officecli close "$FILE"

officecli validate "$FILE"
echo "Generated: $FILE"
