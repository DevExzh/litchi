#!/bin/bash
# Stock Charts Showcase — generates charts-stock.xlsx exercising the xlsx
# `chart` element with chartType=stock: OHLC series, hi-low lines, up-down
# bars, axis/title/legend styling, data labels, reference lines, borders.
#
# CLI twin of charts-stock.py (officecli Python SDK). Both produce an
# equivalent charts-stock.xlsx.
#
# Usage: ./charts-stock.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-stock.xlsx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Sheet: 1-Stock Fundamentals
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="1-Stock Fundamentals"

# Chart 1: Basic OHLC stock chart
# Features: chartType=stock, 4 series (Open/High/Low/Close), catTitle, axisTitle
officecli add "$FILE" "/1-Stock Fundamentals" --type chart \
  --prop chartType=stock \
  --prop title="ACME Corp Weekly OHLC" \
  --prop series1=Open:142,145,148,150,147,152 \
  --prop series2=High:148,151,155,156,153,158 \
  --prop series3=Low:139,142,145,147,144,149 \
  --prop series4=Close:145,148,150,147,152,155 \
  --prop "categories=Week 1,Week 2,Week 3,Week 4,Week 5,Week 6" \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop catTitle=Week --prop "axisTitle=Price (\$)" \
  --prop legend=bottom

# Chart 2: Stock with gridlines and axisfont
# Features: gridlines, axisfont on stock chart
officecli add "$FILE" "/1-Stock Fundamentals" --type chart \
  --prop chartType=stock \
  --prop title="Tech Sector Daily" \
  --prop series1=Open:210,215,212,218,220 \
  --prop series2=High:218,222,219,225,228 \
  --prop series3=Low:207,211,208,214,216 \
  --prop series4=Close:215,212,218,220,225 \
  --prop categories=Mon,Tue,Wed,Thu,Fri \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop gridlines=D9D9D9:0.5 \
  --prop axisfont=9:666666 \
  --prop legend=bottom

# Chart 3: Stock with hiLowLines
# Features: hiLowLines=true (vertical lines connecting high to low)
officecli add "$FILE" "/1-Stock Fundamentals" --type chart \
  --prop chartType=stock \
  --prop title="Energy Sector with Hi-Low Lines" \
  --prop series1=Open:78,80,82,79,83,85 \
  --prop series2=High:84,86,88,85,89,91 \
  --prop series3=Low:75,77,79,76,80,82 \
  --prop series4=Close:80,82,79,83,85,88 \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop hiLowLines=true \
  --prop legend=bottom

# Chart 4: Stock with updownbars
# Features: updownbars=gapWidth:upColor:downColor
officecli add "$FILE" "/1-Stock Fundamentals" --type chart \
  --prop chartType=stock \
  --prop title="Pharma Index with Up-Down Bars" \
  --prop series1=Open:55,58,56,60,62,59 \
  --prop series2=High:61,63,62,66,68,65 \
  --prop series3=Low:52,55,53,57,59,56 \
  --prop series4=Close:58,56,60,62,59,63 \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop updownbars=100:70AD47:C00000 \
  --prop legend=bottom

# ==========================================================================
# Sheet: 2-Stock Styling
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="2-Stock Styling"

# Chart 1: Title styling, legend positioning
# Features: title.font/size/color/bold, legend=right, legendfont
officecli add "$FILE" "/2-Stock Styling" --type chart \
  --prop chartType=stock \
  --prop title="Styled Stock Chart" \
  --prop series1=Open:165,170,168,172,175 \
  --prop series2=High:175,178,176,180,183 \
  --prop series3=Low:160,165,163,168,170 \
  --prop series4=Close:170,168,172,175,180 \
  --prop categories=Mon,Tue,Wed,Thu,Fri \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop legend=right --prop legendfont=10:333333:Calibri

# Chart 2: Series effects, axisLine, catAxisLine
# Features: axisLine, catAxisLine on stock chart
officecli add "$FILE" "/2-Stock Styling" --type chart \
  --prop chartType=stock \
  --prop title="Axis Line Styling" \
  --prop series1=Open:92,95,93,97,99 \
  --prop series2=High:99,102,100,104,106 \
  --prop series3=Low:88,91,89,93,95 \
  --prop series4=Close:95,93,97,99,103 \
  --prop categories=W1,W2,W3,W4,W5 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop hiLowLines=true \
  --prop axisLine=333333:1.5 --prop catAxisLine=333333:1.5 \
  --prop legend=bottom

# Chart 3: axisMin/Max, majorUnit
# Features: axisMin/Max, majorUnit
officecli add "$FILE" "/2-Stock Styling" --type chart \
  --prop chartType=stock \
  --prop title="Custom Axis Range" \
  --prop series1=Open:120,125,122,128,130 \
  --prop series2=High:132,138,135,140,142 \
  --prop series3=Low:115,120,118,124,126 \
  --prop series4=Close:125,122,128,130,135 \
  --prop "categories=Day 1,Day 2,Day 3,Day 4,Day 5" \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop axisMin=110 --prop axisMax=150 \
  --prop majorUnit=10 \
  --prop updownbars=100:70AD47:C00000 \
  --prop legend=bottom

# Chart 4: plotFill, chartFill, roundedCorners
# Features: plotFill, chartFill, roundedCorners
officecli add "$FILE" "/2-Stock Styling" --type chart \
  --prop chartType=stock \
  --prop title="Styled Chart Area" \
  --prop series1=Open:48,50,52,49,53 \
  --prop series2=High:55,57,59,56,60 \
  --prop series3=Low:44,46,48,45,49 \
  --prop series4=Close:50,52,49,53,56 \
  --prop categories=Mon,Tue,Wed,Thu,Fri \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop plotFill=F0F4F8 --prop chartFill=FAFAFA \
  --prop roundedCorners=true \
  --prop hiLowLines=true \
  --prop legend=bottom

# ==========================================================================
# Sheet: 3-Stock Advanced
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="3-Stock Advanced"

# Chart 1: dataLabels, labelFont
# Features: dataLabels, labelPos, labelFont on stock
officecli add "$FILE" "/3-Stock Advanced" --type chart \
  --prop chartType=stock \
  --prop title="Stock with Data Labels" \
  --prop series1=Open:185,190,188,192,195 \
  --prop series2=High:195,198,196,200,203 \
  --prop series3=Low:180,185,183,188,190 \
  --prop series4=Close:190,188,192,195,200 \
  --prop categories=W1,W2,W3,W4,W5 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=top \
  --prop labelFont=8:666666:false \
  --prop legend=bottom

# Chart 2: referenceLine (support/resistance)
# Features: referenceLine as support/resistance level
officecli add "$FILE" "/3-Stock Advanced" --type chart \
  --prop chartType=stock \
  --prop title="Support & Resistance" \
  --prop series1=Open:105,108,106,110,112,109 \
  --prop series2=High:112,115,113,117,119,116 \
  --prop series3=Low:101,104,102,106,108,105 \
  --prop series4=Close:108,106,110,112,109,113 \
  --prop categories=Jan,Feb,Mar,Apr,May,Jun \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop referenceLine=115:C00000:Resistance \
  --prop hiLowLines=true \
  --prop legend=bottom

# Chart 3: chartArea.border, plotArea.border
# Features: chartArea.border, plotArea.border
officecli add "$FILE" "/3-Stock Advanced" --type chart \
  --prop chartType=stock \
  --prop title="Bordered Stock Chart" \
  --prop series1=Open:72,75,73,77,79 \
  --prop series2=High:79,82,80,84,86 \
  --prop series3=Low:68,71,69,73,75 \
  --prop series4=Close:75,73,77,79,83 \
  --prop categories=Mon,Tue,Wed,Thu,Fri \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop chartArea.border=333333:1.5 \
  --prop plotArea.border=999999:0.75 \
  --prop updownbars=100:70AD47:C00000 \
  --prop legend=bottom

# Chart 4: dispUnits, axisNumFmt
# Features: axisNumFmt (dollar format)
officecli add "$FILE" "/3-Stock Advanced" --type chart \
  --prop chartType=stock \
  --prop title="Large Cap Stock" \
  --prop series1=Open:2850,2900,2880,2920,2950 \
  --prop series2=High:2950,2980,2960,3000,3020 \
  --prop series3=Low:2800,2850,2830,2870,2900 \
  --prop series4=Close:2900,2880,2920,2950,2990 \
  --prop categories=Q1,Q2,Q3,Q4,Q5 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop "axisNumFmt=\$#,##0" \
  --prop hiLowLines=true \
  --prop legend=bottom

# Remove blank default Sheet1 (all data is inline)
officecli remove "$FILE" /Sheet1

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
