#!/bin/bash
# Word Charts Showcase — every docx chart family, embedded inline in a document.
# CLI twin of charts.py (officecli Python SDK). Both produce an equivalent
# charts.docx.
#
# In Word a chart is an INLINE DrawingML object anchored on a paragraph. You
# add it with `add <file> /body/p[N] --type chart ...`, giving the chart its
# data inline (data=/series{N}=/categories=) — Word has no worksheet grid to
# pull a dataRange from, so inline data is the idiomatic path here. The chart
# is then addressed document-globally as `/chart[M]` for get/set/query (NOT
# `/body/p[N]/chart[M]`).
#
# Every meaningful docx chart family is demonstrated at least once:
# chart types (column/bar/line/pie/area/scatter/radar/doughnut/stock/combo
# plus extended cx types funnel/treemap/waterfall), titles & title styling,
# legend, data labels, colors & gradients (areafill), axis scaling/styling,
# display units, radar style, rounded corners, markers, and transparency.
#
# 14 charts, each under its own heading paragraph.
#
# Usage:
#   ./charts.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts.docx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# Every chart anchors on its own body paragraph. We keep a running 1-based
# paragraph index ($P) so each chart's anchor path (/body/p[$P]) is explicit.
# Layout per demo: one Heading2 paragraph, then one empty body paragraph that
# hosts the inline chart. Charts are read back via the flat /chart[N] path.
P=0
heading() { officecli add "$FILE" /body --type paragraph --prop text="$1" --prop style=Heading2; P=$((P+1)); }
host()    { officecli add "$FILE" /body --type paragraph --prop text="";   P=$((P+1)); }

# Title + intro (p[1], p[2])
officecli add "$FILE" /body --type paragraph --prop text="Word Charts Showcase" --prop style=Heading1 --prop align=center; P=$((P+1))
officecli add "$FILE" /body --type paragraph --prop text="Each chart below is an inline DrawingML object anchored on its own paragraph. Charts are addressed document-globally as /chart[N]."; P=$((P+1))

# ==========================================================================
# 1. Column — the workhorse. Axis titles, axis scaling, gridlines, colors.
# ==========================================================================
heading "1. Column — axis titles, scaling & gridlines"; host
# Features: chartType=column, data (inline Name:v;Name2:v), categories, colors,
#   catTitle/axisTitle, axisMin/axisMax/axisNumFmt, gridlines, legend, width/height
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=column \
  --prop title="Quarterly Revenue by Region" \
  --prop 'data=East:120,135,148,162;West:110,118,130,145;South:95,108,115,128' \
  --prop categories="Q1,Q2,Q3,Q4" \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop catTitle=Quarter --prop axisTitle="Revenue (K)" \
  --prop axisMin=0 --prop axisMax=200 --prop axisNumFmt="#,##0" \
  --prop gridlines=D9D9D9:0.5:dot \
  --prop legend=bottom \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 2. Bar — horizontal columns with gap width & data labels.
# ==========================================================================
heading "2. Bar — gap width & data labels"; host
# Features: chartType=bar, gapwidth, dataLabels=value, labelPos=outsideEnd, labelfont
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=bar \
  --prop title="Product Units Sold" \
  --prop 'data=Units:320,280,410,190,360' \
  --prop categories="Laptop,Phone,Tablet,Watch,Buds" \
  --prop colors=2E75B6 \
  --prop gapwidth=80 \
  --prop dataLabels=value --prop labelPos=outsideEnd \
  --prop labelfont=9:333333:Calibri \
  --prop legend=none \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 3. Line — markers, smoothing, drop lines.
# ==========================================================================
heading "3. Line — markers, smoothing & drop lines"; host
# Features: chartType=line, marker (symbol:size), smooth, droplines, linewidth
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=line \
  --prop title="Monthly Active Users" \
  --prop 'data=2023:120,180,210,250,280,310;2024:150,220,260,300,340,380' \
  --prop categories="Jan,Feb,Mar,Apr,May,Jun" \
  --prop colors=4472C4,ED7D31 \
  --prop marker=circle:6 \
  --prop smooth=true \
  --prop droplines=808080:0.5 \
  --prop linewidth=2 \
  --prop legend=bottom \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 4. Pie — percent labels, slice explosion, first-slice angle.
# ==========================================================================
heading "4. Pie — percent labels & slice explosion"; host
# Features: chartType=pie, dataLabels=percent, explosion, firstSliceAngle, colors
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=pie \
  --prop title="Market Share" \
  --prop 'data=Share:42,28,18,12' \
  --prop categories="Alpha,Beta,Gamma,Other" \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop dataLabels=percent \
  --prop explosion=8 \
  --prop firstSliceAngle=90 \
  --prop legend=right \
  --prop width=14cm --prop height=9cm

# ==========================================================================
# 5. Area — gradient fill via areafill (docx-specific shortcut).
# ==========================================================================
heading "5. Area — gradient fill (areafill)"; host
# Features: chartType=area, areafill=c1-c2:angle (gradient applied to every series shape)
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=area \
  --prop title="Cumulative Traffic" \
  --prop 'data=Visits:20,35,30,55,48,70,65' \
  --prop categories="Mon,Tue,Wed,Thu,Fri,Sat,Sun" \
  --prop areafill=4472C4-A5C8FF:90 \
  --prop gridlines=E0E0E0:0.5:solid \
  --prop legend=none \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 6. Scatter — XY scatter with smoothed-marker style.
# ==========================================================================
heading "6. Scatter — smoothMarker style"; host
# Features: chartType=scatter, scatterstyle=smoothMarker, marker, catTitle/axisTitle
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=scatter \
  --prop title="Load vs Response Time" \
  --prop series1="Latency:12,18,27,41,60,88" \
  --prop categories="10,20,40,80,160,320" \
  --prop scatterstyle=smoothMarker \
  --prop marker=diamond:7:C00000 \
  --prop catTitle="Concurrent Users" --prop axisTitle="ms" \
  --prop legend=none \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 7. Radar — filled radar style (radarstyle docx shortcut).
# ==========================================================================
heading "7. Radar — filled style (radarstyle)"; host
# Features: chartType=radar, radarstyle=filled, multi-series, transparency
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=radar \
  --prop title="Product Comparison" \
  --prop 'data=Model A:4,5,3,4,5;Model B:5,3,4,5,3' \
  --prop categories="Speed,Battery,Camera,Price,Display" \
  --prop radarstyle=filled \
  --prop colors=4472C4,ED7D31 \
  --prop transparency=40 \
  --prop legend=bottom \
  --prop width=14cm --prop height=10cm

# ==========================================================================
# 8. Doughnut — hole size & percent labels.
# ==========================================================================
heading "8. Doughnut — hole size & percent labels"; host
# Features: chartType=doughnut, holeSize, dataLabels=percent, colors
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=doughnut \
  --prop title="Budget Allocation" \
  --prop 'data=Budget:35,25,20,20' \
  --prop categories="R&D,Sales,Ops,Admin" \
  --prop holeSize=55 \
  --prop dataLabels=percent \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop legend=right \
  --prop width=14cm --prop height=9cm

# ==========================================================================
# 9. Stock — high/low/close (OHLC-style) series.
# ==========================================================================
heading "9. Stock — high / low / close series"; host
# Features: chartType=stock, three ordered series (High, Low, Close), hilowlines
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=stock \
  --prop title="Share Price (5 Days)" \
  --prop series1="High:32,35,34,38,37" \
  --prop series2="Low:28,29,30,32,33" \
  --prop series3="Close:30,34,31,37,35" \
  --prop categories="Mon,Tue,Wed,Thu,Fri" \
  --prop hilowlines=true \
  --prop legend=bottom \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 10. Combo — mixed column + line on a secondary axis.
# ==========================================================================
heading "10. Combo — column + line on secondary axis"; host
# Features: chartType=combo, combotypes=column,line, secondaryaxis=2
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=combo \
  --prop title="Revenue vs Growth Rate" \
  --prop series1="Revenue:120,180,250,310,380" \
  --prop series2="Growth %:50,33,39,24,23" \
  --prop categories="2021,2022,2023,2024,2025" \
  --prop combotypes="column,line" \
  --prop secondaryaxis=2 \
  --prop colors=2E75B6,C00000 \
  --prop legend=bottom \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 11. Column with display units & rounded chart corners + title styling.
# ==========================================================================
heading "11. Column — display units & rounded corners"; host
# Features: dispUnits=thousands, roundedcorners=true, chartFill, title.font/size/color/bold
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=column \
  --prop title="Revenue (in Thousands)" \
  --prop 'data=Revenue:12000,18500,22000,31000,45000' \
  --prop categories="2021,2022,2023,2024,2025" \
  --prop colors=1F4E79 \
  --prop dispUnits=thousands \
  --prop axisNumFmt="#,##0" \
  --prop roundedcorners=true \
  --prop chartFill=F8F8F8 \
  --prop title.font=Georgia --prop title.size=15 --prop title.color=1F4E79 --prop title.bold=true \
  --prop legend=none \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 12. Funnel — extended cx chart type.
# ==========================================================================
heading "12. Funnel — extended (cx) chart"; host
# Features: chartType=funnel (extended cx:chart), single-series stage breakdown
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=funnel \
  --prop title="Sales Funnel" \
  --prop 'data=Stage:1000,720,430,210,95' \
  --prop categories="Visitors,Leads,MQL,SQL,Won" \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 13. Treemap — extended cx chart type.
# ==========================================================================
heading "13. Treemap — extended (cx) chart"; host
# Features: chartType=treemap (extended cx:chart), proportional area tiles
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=treemap \
  --prop title="Storage by File Type" \
  --prop 'data=Size:420,310,180,90,45' \
  --prop categories="Video,Images,Docs,Audio,Other" \
  --prop width=16cm --prop height=9cm

# ==========================================================================
# 14. Waterfall — extended cx chart with increase/decrease/total colors.
# ==========================================================================
heading "14. Waterfall — increase / decrease / total colors"; host
# Features: chartType=waterfall, increaseColor, decreaseColor, totalColor (all add-time)
officecli add "$FILE" "/body/p[$P]" --type chart \
  --prop chartType=waterfall \
  --prop title="Cash Flow" \
  --prop 'data=Cash:100,-30,50,-20,80' \
  --prop categories="Start,Q1,Q2,Q3,End" \
  --prop increaseColor=00AA00 \
  --prop decreaseColor=C00000 \
  --prop totalColor=4472C4 \
  --prop width=16cm --prop height=9cm

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
