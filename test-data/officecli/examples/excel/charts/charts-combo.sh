#!/bin/bash
# Combo Charts Showcase — column+line, column+area, secondary axes, and styling.
# Generates: charts-combo.xlsx (16 combo charts across 4 sheets)
# CLI twin of charts-combo.py (officecli Python SDK).
# Usage: ./charts-combo.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-combo.xlsx"

rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Sheet: 1-Combo Fundamentals
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="1-Combo Fundamentals"

# Chart 1: Basic combo with comboSplit (2 bar series + 1 line)
# Features: chartType=combo, comboSplit=2 (first 2 as bars, rest as lines)
officecli add "$FILE" "/1-Combo Fundamentals" --type chart \
  --prop chartType=combo \
  --prop title="Revenue vs Expenses vs Margin" \
  --prop series1="Revenue:120,145,160,180,195" \
  --prop series2="Expenses:90,100,110,115,125" \
  --prop series3="Margin %:25,31,31,36,36" \
  --prop categories=Q1,Q2,Q3,Q4,Q5 \
  --prop comboSplit=2 \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop legend=bottom

# Chart 2: Combo with secondaryAxis (line on right Y-axis)
# Features: secondaryAxis=2 (series 2 on right Y-axis), catTitle, axisTitle
officecli add "$FILE" "/1-Combo Fundamentals" --type chart \
  --prop chartType=combo \
  --prop title="Sales & Growth Rate" \
  --prop series1="Sales ($K):320,380,420,510,560" \
  --prop series2="Growth %:8,19,11,21,10" \
  --prop categories=2021,2022,2023,2024,2025 \
  --prop comboSplit=1 \
  --prop secondaryAxis=2 \
  --prop colors=2E75B6,C00000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop legend=bottom \
  --prop catTitle=Year --prop axisTitle="Sales ($K)"

# Chart 3: combotypes per-series type control
# Features: combotypes=column,column,line,area (per-series type)
officecli add "$FILE" "/1-Combo Fundamentals" --type chart \
  --prop chartType=combo \
  --prop title="Mixed Series Types" \
  --prop series1="Product A:50,65,70,80,90" \
  --prop series2="Product B:40,55,60,72,85" \
  --prop series3="Trend:48,62,68,78,88" \
  --prop series4="Forecast:30,40,50,55,65" \
  --prop categories=Jan,Feb,Mar,Apr,May \
  --prop combotypes=column,column,line,area \
  --prop colors=4472C4,ED7D31,70AD47,BDD7EE \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom

# Chart 4: combotypes with secondaryAxis
# Features: combotypes + secondaryAxis together
officecli add "$FILE" "/1-Combo Fundamentals" --type chart \
  --prop chartType=combo \
  --prop title="Revenue Mix & Margin" \
  --prop series1="Domestic:200,220,250,270,300" \
  --prop series2="Export:80,95,110,130,150" \
  --prop series3="Net Margin %:18,20,22,24,26" \
  --prop categories=2021,2022,2023,2024,2025 \
  --prop combotypes=column,column,line \
  --prop secondaryAxis=3 \
  --prop colors=4472C4,9DC3E6,C00000 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom \
  --prop catTitle=Year

# ==========================================================================
# Sheet: 2-Combo Styling
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="2-Combo Styling"

# Chart 1: Title, legend, axisfont styling
# Features: title.font/size/color/bold, legendfont, axisfont
officecli add "$FILE" "/2-Combo Styling" --type chart \
  --prop chartType=combo \
  --prop title="Styled Combo Chart" \
  --prop series1="Revenue:150,175,200,220" \
  --prop series2="COGS:100,110,130,140" \
  --prop series3="Profit %:33,37,35,36" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop comboSplit=2 \
  --prop colors=1F4E79,5B9BD5,70AD47 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop legend=bottom --prop legendfont=10:333333:Calibri \
  --prop axisfont=9:666666

# Chart 2: Series shadow, gradients
# Features: gradients (per-bar-series), series.shadow
officecli add "$FILE" "/2-Combo Styling" --type chart \
  --prop chartType=combo \
  --prop title="Gradient & Shadow Effects" \
  --prop series1="Actual:85,92,105,120,135" \
  --prop series2="Budget:80,90,100,110,120" \
  --prop series3="Variance:5,2,5,10,15" \
  --prop categories=Jan,Feb,Mar,Apr,May \
  --prop comboSplit=2 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop gradients="1F4E79-5B9BD5:90;C55A11-F4B183:90" \
  --prop series.shadow=000000-4-315-2-30 \
  --prop legend=bottom

# Chart 3: dataLabels on line series
# Features: dataLabels=true, labelPos=top, labelFont
officecli add "$FILE" "/2-Combo Styling" --type chart \
  --prop chartType=combo \
  --prop title="Data Labels on Lines" \
  --prop series1="Units:500,620,710,800" \
  --prop series2="Avg Price:45,48,52,55" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop comboSplit=1 \
  --prop secondaryAxis=2 \
  --prop colors=4472C4,ED7D31 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=top \
  --prop labelFont=9:333333:true \
  --prop legend=bottom

# Chart 4: plotFill, chartFill, roundedCorners
# Features: plotFill, chartFill, roundedCorners
officecli add "$FILE" "/2-Combo Styling" --type chart \
  --prop chartType=combo \
  --prop title="Chart Area Styling" \
  --prop series1="Online:180,210,240,260,290" \
  --prop series2="Retail:150,140,135,130,120" \
  --prop series3="Growth %:5,12,15,10,12" \
  --prop categories=2021,2022,2023,2024,2025 \
  --prop comboSplit=2 \
  --prop colors=2E75B6,ED7D31,70AD47 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop plotFill=F0F4F8 --prop chartFill=FAFAFA \
  --prop roundedCorners=true \
  --prop legend=bottom

# ==========================================================================
# Sheet: 3-Combo Advanced
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="3-Combo Advanced"

# Chart 1: referenceLine, gridlines
# Features: referenceLine=value:label:color, gridlines
officecli add "$FILE" "/3-Combo Advanced" --type chart \
  --prop chartType=combo \
  --prop title="Target Reference Line" \
  --prop series1="Actual:95,105,115,125,130" \
  --prop series2="Forecast:90,100,110,120,130" \
  --prop categories=Jan,Feb,Mar,Apr,May \
  --prop comboSplit=1 \
  --prop colors=4472C4,BDD7EE \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop referenceLine=110:C00000:Target \
  --prop gridlines=D9D9D9:0.5 \
  --prop legend=bottom

# Chart 2: axisMin/Max, dispUnits
# Features: axisMin/Max, dispUnits=thousands
officecli add "$FILE" "/3-Combo Advanced" --type chart \
  --prop chartType=combo \
  --prop title="Axis Scaling & Units" \
  --prop series1="Revenue:1200000,1450000,1600000,1800000" \
  --prop series2="Profit %:18,22,25,28" \
  --prop categories=2022,2023,2024,2025 \
  --prop comboSplit=1 \
  --prop secondaryAxis=2 \
  --prop colors=2E75B6,70AD47 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop axisMin=1000000 --prop axisMax=2000000 \
  --prop dispUnits=thousands \
  --prop legend=bottom

# Chart 3: Manual layout
# Features: plotLayout=left,top,width,height (manual plot area)
officecli add "$FILE" "/3-Combo Advanced" --type chart \
  --prop chartType=combo \
  --prop title="Manual Layout" \
  --prop series1="Plan:100,120,140,160" \
  --prop series2="Actual:95,125,135,170" \
  --prop series3="Delta %:-5,4,-4,6" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop comboSplit=2 \
  --prop secondaryAxis=3 \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop plotLayout=0.1,0.15,0.85,0.75 \
  --prop legend=bottom

# Chart 4: Multiple line series with markers + bar series
# Features: multiple line series on secondary axis, markers
officecli add "$FILE" "/3-Combo Advanced" --type chart \
  --prop chartType=combo \
  --prop title="Multi-Line with Markers" \
  --prop series1="Units Sold:800,920,1050,1200,1350" \
  --prop series2="North:30,35,38,42,45" \
  --prop series3="South:25,28,32,36,40" \
  --prop series4="West:20,24,28,32,35" \
  --prop categories=Q1,Q2,Q3,Q4,Q5 \
  --prop comboSplit=1 \
  --prop secondaryAxis=2,3,4 \
  --prop colors=4472C4,C00000,70AD47,FFC000 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop markers=circle-6 \
  --prop legend=bottom

# ==========================================================================
# Sheet: 4-Combo Effects
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="4-Combo Effects"

# Chart 1: title.glow, title.shadow
# Features: title.glow=color-radius, title.shadow
officecli add "$FILE" "/4-Combo Effects" --type chart \
  --prop chartType=combo \
  --prop title="Glowing Title" \
  --prop series1="Metric A:60,72,85,90,100" \
  --prop series2="Metric B:40,50,55,62,70" \
  --prop series3="Ratio:67,69,65,69,70" \
  --prop categories=W1,W2,W3,W4,W5 \
  --prop comboSplit=2 \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop title.glow=4472C4-6 \
  --prop title.shadow=000000-3-315-2-30 \
  --prop legend=bottom

# Chart 2: chartArea.border, plotArea.border
# Features: chartArea.border=color-width, plotArea.border
officecli add "$FILE" "/4-Combo Effects" --type chart \
  --prop chartType=combo \
  --prop title="Bordered Areas" \
  --prop series1="Income:250,280,310,340" \
  --prop series2="Costs:180,195,210,225" \
  --prop series3="Margin %:28,30,32,34" \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop comboSplit=2 \
  --prop colors=2E75B6,ED7D31,548235 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop chartArea.border=333333:1.5 \
  --prop plotArea.border=999999:0.75 \
  --prop legend=bottom

# Chart 3: colorRule
# Features: colorRule=threshold:belowColor:aboveColor
officecli add "$FILE" "/4-Combo Effects" --type chart \
  --prop chartType=combo \
  --prop title="Color Rule Combo" \
  --prop series1="Performance:72,85,65,90,78" \
  --prop series2="Target:80,80,80,80,80" \
  --prop categories="Team A,Team B,Team C,Team D,Team E" \
  --prop comboSplit=1 \
  --prop colors=4472C4,C00000 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop colorRule=80:C00000:70AD47 \
  --prop legend=bottom

# Chart 4: Complex combo with 5+ series
# Features: 5 series, mixed combotypes, secondary axis
officecli add "$FILE" "/4-Combo Effects" --type chart \
  --prop chartType=combo \
  --prop title="Full Business Dashboard" \
  --prop series1="Revenue:500,550,600,650,700" \
  --prop series2="COGS:300,320,340,360,380" \
  --prop series3="OpEx:100,105,110,115,120" \
  --prop series4="Net Income:100,125,150,175,200" \
  --prop series5="Margin %:20,23,25,27,29" \
  --prop categories=2021,2022,2023,2024,2025 \
  --prop combotypes=column,column,column,area,line \
  --prop secondaryAxis=5 \
  --prop colors=4472C4,ED7D31,A5A5A5,BDD7EE,C00000 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom \
  --prop gridlines=E0E0E0:0.5

# Remove blank default Sheet1 (all data is inline)
officecli remove "$FILE" /Sheet1

officecli close "$FILE"

officecli validate "$FILE"
echo "Generated: $FILE"
