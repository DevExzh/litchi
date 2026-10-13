#!/bin/bash
# Bar (Horizontal) Charts Showcase — generates charts-bar.xlsx exercising the
# full xlsx bar/barStacked/barPercentStacked/bar3d chart family.
#
# CLI twin of charts-bar.py (officecli Python SDK). Both produce an equivalent
# charts-bar.xlsx.
#
# 8 sheets (Sheet1 data + 7 chart sheets), 28 charts total.
#
# Usage: ./charts-bar.sh

FILE="$(dirname "$0")/charts-bar.xlsx"
rm -f "$FILE"

# Forward-compat tolerance: this showcase deliberately exercises a few props the
# chart handler doesn't consume yet (e.g. point{N}.color → exit 2
# unsupported_property) and a few values it rejects on some chart types (e.g.
# axisposition=nextTo). officecli warns but still creates the element; we log the
# warning and press on rather than aborting, so the whole 28-chart showcase
# runs — matching the SDK twin, whose doc.batch() doesn't abort on these either.
officecli() {
  command officecli "$@" || echo "  (officecli exit $? — continuing; prop unsupported on this chart type)"
}

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Source data — shared across all charts
# ==========================================================================
echo "--- Populating source data ---"
officecli set "$FILE" /Sheet1/A1 --prop text=Department --prop bold=true
officecli set "$FILE" /Sheet1/B1 --prop text=Q1 --prop bold=true
officecli set "$FILE" /Sheet1/C1 --prop text=Q2 --prop bold=true
officecli set "$FILE" /Sheet1/D1 --prop text=Q3 --prop bold=true
officecli set "$FILE" /Sheet1/E1 --prop text=Q4 --prop bold=true

officecli set "$FILE" /Sheet1/A2 --prop text=Engineering
officecli set "$FILE" /Sheet1/B2 --prop text=185
officecli set "$FILE" /Sheet1/C2 --prop text=195
officecli set "$FILE" /Sheet1/D2 --prop text=210
officecli set "$FILE" /Sheet1/E2 --prop text=228
officecli set "$FILE" /Sheet1/A3 --prop text=Marketing
officecli set "$FILE" /Sheet1/B3 --prop text=120
officecli set "$FILE" /Sheet1/C3 --prop text=135
officecli set "$FILE" /Sheet1/D3 --prop text=142
officecli set "$FILE" /Sheet1/E3 --prop text=158
officecli set "$FILE" /Sheet1/A4 --prop text=Sales
officecli set "$FILE" /Sheet1/B4 --prop text=210
officecli set "$FILE" /Sheet1/C4 --prop text=225
officecli set "$FILE" /Sheet1/D4 --prop text=240
officecli set "$FILE" /Sheet1/E4 --prop text=260
officecli set "$FILE" /Sheet1/A5 --prop text=Support
officecli set "$FILE" /Sheet1/B5 --prop text=95
officecli set "$FILE" /Sheet1/C5 --prop text=105
officecli set "$FILE" /Sheet1/D5 --prop text=112
officecli set "$FILE" /Sheet1/E5 --prop text=118
officecli set "$FILE" /Sheet1/A6 --prop text=Finance
officecli set "$FILE" /Sheet1/B6 --prop text=78
officecli set "$FILE" /Sheet1/C6 --prop text=82
officecli set "$FILE" /Sheet1/D6 --prop text=88
officecli set "$FILE" /Sheet1/E6 --prop text=92
officecli set "$FILE" /Sheet1/A7 --prop text=HR
officecli set "$FILE" /Sheet1/B7 --prop text=62
officecli set "$FILE" /Sheet1/C7 --prop text=68
officecli set "$FILE" /Sheet1/D7 --prop text=72
officecli set "$FILE" /Sheet1/E7 --prop text=78
officecli set "$FILE" /Sheet1/A8 --prop text=Legal
officecli set "$FILE" /Sheet1/B8 --prop text=55
officecli set "$FILE" /Sheet1/C8 --prop text=58
officecli set "$FILE" /Sheet1/D8 --prop text=62
officecli set "$FILE" /Sheet1/E8 --prop text=68
officecli set "$FILE" /Sheet1/A9 --prop text=Operations
officecli set "$FILE" /Sheet1/B9 --prop text=140
officecli set "$FILE" /Sheet1/C9 --prop text=152
officecli set "$FILE" /Sheet1/D9 --prop text=165
officecli set "$FILE" /Sheet1/E9 --prop text=178

# ==========================================================================
# Sheet: 1-Bar Fundamentals
# ==========================================================================
echo "--- 1-Bar Fundamentals ---"
officecli add "$FILE" / --type sheet --prop name="1-Bar Fundamentals"

# Chart 1: Basic bar chart with dataRange, axis titles, and gridlines
# Features: chartType=bar, dataRange, catTitle, axisTitle, axisfont, gridlines
officecli add "$FILE" "/1-Bar Fundamentals" --type chart \
  --prop chartType=bar \
  --prop title="Department Performance — Q1" \
  --prop dataRange=Sheet1!A1:B9 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop catTitle=Department --prop axisTitle=Score \
  --prop axisfont=9:333333:Arial \
  --prop gridlines=D9D9D9:0.5:dot

# Chart 2: Inline series with custom colors, gap width, and data labels
# Features: inline series, colors per category, gapwidth, dataLabels=outsideEnd
officecli add "$FILE" "/1-Bar Fundamentals" --type chart \
  --prop chartType=bar \
  --prop title="Survey Results" \
  --prop series1=Satisfaction:85,72,91,68,78 \
  --prop categories=Product,Service,Delivery,Price,Overall \
  --prop colors=4472C4,ED7D31,70AD47,FFC000,5B9BD5 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop gapwidth=80 \
  --prop dataLabels=outsideEnd

# Chart 3: Stacked bar with overlap and series outline
# Features: barStacked, overlap=0, series.outline (white separator)
officecli add "$FILE" "/1-Bar Fundamentals" --type chart \
  --prop chartType=barStacked \
  --prop title="Quarterly Headcount by Dept" \
  --prop series1=Q1:30,18,25,12 \
  --prop series2=Q2:35,20,28,14 \
  --prop series3=Q3:38,22,30,16 \
  --prop categories=Engineering,Marketing,Sales,Support \
  --prop colors=2E75B6,70AD47,FFC000 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop overlap=0 \
  --prop series.outline=FFFFFF-0.5

# Chart 4: data= shorthand with legend=bottom
# Features: data= shorthand (inline multi-series), legend=bottom
officecli add "$FILE" "/1-Bar Fundamentals" --type chart \
  --prop chartType=bar \
  --prop title="Training Hours by Team" \
  --prop "data=Technical:45,38,52;Soft Skills:20,28,18;Compliance:12,15,10" \
  --prop categories=Engineering,Sales,Support \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom

# ==========================================================================
# Sheet: 2-Bar Variants
# ==========================================================================
echo "--- 2-Bar Variants ---"
officecli add "$FILE" / --type sheet --prop name="2-Bar Variants"

# Chart 1: barStacked with tight gap width
# Features: barStacked, gapwidth=50 (tight bars)
officecli add "$FILE" "/2-Bar Variants" --type chart \
  --prop chartType=barStacked \
  --prop title="Budget Allocation" \
  --prop series1=Salaries:120,80,95,60 \
  --prop series2=Operations:45,35,40,25 \
  --prop series3=Marketing:30,50,20,15 \
  --prop categories=Engineering,Sales,Support,HR \
  --prop colors=1F4E79,2E75B6,9DC3E6 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop gapwidth=50 \
  --prop legend=bottom

# Chart 2: barPercentStacked with axis number format and reference line
# Features: barPercentStacked, axisNumFmt=0%, referenceLine with label and dash
officecli add "$FILE" "/2-Bar Variants" --type chart \
  --prop chartType=barPercentStacked \
  --prop title="Task Completion Ratio" \
  --prop series1=Done:75,60,90,45,80 \
  --prop "series2=In Progress:15,25,5,30,12" \
  --prop series3=Blocked:10,15,5,25,8 \
  --prop categories=Backend,Frontend,QA,Design,DevOps \
  --prop colors=70AD47,FFC000,C00000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop axisNumFmt=0% \
  --prop referenceLine=0.5:FF0000:Target:dash \
  --prop legend=bottom

# Chart 3: bar3d with perspective and style
# Features: bar3d, view3d (rotX,rotY,perspective), style=3
officecli add "$FILE" "/2-Bar Variants" --type chart \
  --prop chartType=bar3d \
  --prop title="3D Revenue by Region" \
  --prop series1=Revenue:340,280,310,195 \
  --prop categories=North,South,East,West \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=10,30,20 \
  --prop style=3 \
  --prop legend=right

# Chart 4: bar3d with cylinder shape
# Features: bar3d shape=cylinder, multi-series 3D bars
officecli add "$FILE" "/2-Bar Variants" --type chart \
  --prop chartType=bar3d \
  --prop title="Cylinder — Project Milestones" \
  --prop series1=Completed:8,12,6,10,15 \
  --prop series2=Remaining:4,3,6,5,2 \
  --prop categories=Alpha,Beta,Gamma,Delta,Epsilon \
  --prop colors=2E75B6,BDD7EE \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop shape=cylinder \
  --prop gapwidth=60 \
  --prop legend=bottom

# ==========================================================================
# Sheet: 3-Bar Styling
# ==========================================================================
echo "--- 3-Bar Styling ---"
officecli add "$FILE" / --type sheet --prop name="3-Bar Styling"

# Chart 1: Title styling (font, size, color, bold)
# Features: title.font, title.size, title.color, title.bold
officecli add "$FILE" "/3-Bar Styling" --type chart \
  --prop chartType=bar \
  --prop title="Styled Title Demo" \
  --prop series1=Score:88,76,92,65,84 \
  --prop categories=Dept A,Dept B,Dept C,Dept D,Dept E \
  --prop colors=4472C4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop gapwidth=100

# Chart 2: Series shadow and outline effects
# Features: series.shadow (color-blur-angle-dist-opacity), series.outline
officecli add "$FILE" "/3-Bar Styling" --type chart \
  --prop chartType=bar \
  --prop title="Shadow & Outline" \
  --prop series1=2024:165,142,180,128 \
  --prop series2=2025:185,158,195,140 \
  --prop categories=Engineering,Marketing,Sales,Support \
  --prop colors=2E75B6,ED7D31 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop series.shadow=000000-4-315-2-30 \
  --prop series.outline=1F4E79-1 \
  --prop legend=bottom

# Chart 3: Per-series gradients
# Features: gradients (per-bar gradient fills, angle=0 for horizontal), labelFont (size:color:bold)
officecli add "$FILE" "/3-Bar Styling" --type chart \
  --prop chartType=bar \
  --prop title="Gradient Bars" \
  --prop series1=Revenue:320,275,410,190,245 \
  --prop categories=North,South,East,West,Central \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop "gradients=1F4E79-5B9BD5:0;C55A11-F4B183:0;548235-A9D18E:0;7F6000-FFD966:0;843C0B-DDA15E:0" \
  --prop dataLabels=outsideEnd \
  --prop labelFont=9:333333:true

# Chart 4: Plot fill gradient, chart fill, transparency, rounded corners
# Features: plotFill gradient, chartFill, transparency, roundedCorners
officecli add "$FILE" "/3-Bar Styling" --type chart \
  --prop chartType=bar \
  --prop title="Styled Background" \
  --prop dataRange=Sheet1!A1:C9 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=5B9BD5,ED7D31 \
  --prop plotFill=F0F4F8-D6E4F0:90 \
  --prop chartFill=FFFFFF \
  --prop transparency=20 \
  --prop roundedCorners=true \
  --prop legend=right

# ==========================================================================
# Sheet: 4-Axis & Labels
# ==========================================================================
echo "--- 4-Axis & Labels ---"
officecli add "$FILE" / --type sheet --prop name="4-Axis & Labels"

# Chart 1: Custom axis min/max, majorUnit, and gridlines styling
# Features: axisMin, axisMax, majorUnit, gridlines styling, minorGridlines, axisLine, catAxisLine
officecli add "$FILE" "/4-Axis & Labels" --type chart \
  --prop chartType=bar \
  --prop title="Axis Scale (50–250)" \
  --prop dataRange=Sheet1!A1:B9 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop axisMin=50 --prop axisMax=250 --prop majorUnit=50 \
  --prop gridlines=D0D0D0:0.5:solid \
  --prop minorGridlines=EEEEEE:0.3:dot \
  --prop axisLine=C00000:1.5:solid \
  --prop catAxisLine=2E75B6:1.5:solid

# Chart 2: Log scale, axis reverse, and display units
# Features: logBase=10, axisReverse=true, dispUnits=thousands
officecli add "$FILE" "/4-Axis & Labels" --type chart \
  --prop chartType=bar \
  --prop title="Log Scale & Reverse" \
  --prop "series1=Users:10,100,1000,5000,25000,100000" \
  --prop "categories=Tier 1,Tier 2,Tier 3,Tier 4,Tier 5,Tier 6" \
  --prop colors=2E75B6 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop logBase=10 \
  --prop axisReverse=true \
  --prop dispUnits=thousands \
  --prop gridlines=E0E0E0:0.5:dash

# Chart 3: Data labels with labelFont, numFmt, separator
# Features: dataLabels, labelFont, dataLabels.numFmt, dataLabels.separator
officecli add "$FILE" "/4-Axis & Labels" --type chart \
  --prop chartType=bar \
  --prop title="Labeled Metrics" \
  --prop series1=FY2025:148,92,215,178,125 \
  --prop categories=Revenue,Costs,Gross,EBITDA,Net Income \
  --prop colors=4472C4 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop dataLabels=outsideEnd \
  --prop labelFont=10:1F4E79:true \
  --prop dataLabels.numFmt=#,##0 \
  --prop "dataLabels.separator=: "

# Chart 4: Per-point label delete/text and per-point color
# Features: dataLabel{N}.delete, dataLabel{N}.text, point{N}.color
officecli add "$FILE" "/4-Axis & Labels" --type chart \
  --prop chartType=bar \
  --prop title="Highlight Winner" \
  --prop series1=Score:72,85,68,95,78 \
  --prop categories=Team A,Team B,Team C,Team D,Team E \
  --prop colors=9DC3E6 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop dataLabel1.delete=true --prop dataLabel3.delete=true \
  --prop dataLabel5.delete=true \
  --prop dataLabel4.text="Winner!" \
  --prop point4.color=C00000 \
  --prop point2.color=2E75B6 \
  --prop gapwidth=70

# ==========================================================================
# Sheet: 5-Legend & Layout
# ==========================================================================
echo "--- 5-Legend & Layout ---"
officecli add "$FILE" / --type sheet --prop name="5-Legend & Layout"

# Chart 1: Legend positions (right)
# Features: legend=right (4-series bar with legend on right)
officecli add "$FILE" "/5-Legend & Layout" --type chart \
  --prop chartType=bar \
  --prop title="Legend: Right" \
  --prop dataRange=Sheet1!A1:E9 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop legend=right

# Chart 2: Legend font styling and overlay
# Features: legendfont (size:color:fontname), legend.overlay=true
officecli add "$FILE" "/5-Legend & Layout" --type chart \
  --prop chartType=bar \
  --prop title="Legend: Font & Overlay" \
  --prop dataRange=Sheet1!A1:E9 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colors=1F4E79,2E75B6,5B9BD5,9DC3E6 \
  --prop legend=top \
  --prop legend.overlay=true \
  --prop legendfont=10:1F4E79:Calibri

# Chart 3: Manual layout — plotArea.x/y/w/h, title.x/y, legend.x/y/w/h
# Features: plotArea.x/y/w/h, title.x/y, legend.x/y/w/h (manual layout)
officecli add "$FILE" "/5-Legend & Layout" --type chart \
  --prop chartType=bar \
  --prop title="Manual Layout" \
  --prop dataRange=Sheet1!A1:C9 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=2E75B6,70AD47 \
  --prop plotArea.x=0.25 --prop plotArea.y=0.15 \
  --prop plotArea.w=0.70 --prop plotArea.h=0.60 \
  --prop title.x=0.20 --prop title.y=0.02 \
  --prop legend.x=0.25 --prop legend.y=0.82 \
  --prop legend.w=0.50 --prop legend.h=0.10 \
  --prop title.font=Arial --prop title.size=13 \
  --prop title.bold=true

# Chart 4: Secondary axis with chart/plot area borders
# Features: secondaryAxis=2, chartArea.border, plotArea.border
officecli add "$FILE" "/5-Legend & Layout" --type chart \
  --prop chartType=bar \
  --prop title="Dual Axis: Revenue vs Margin" \
  --prop "series1=Revenue:340,280,410,195,310" \
  --prop "series2=Margin %:22,18,28,15,25" \
  --prop categories=North,South,East,West,Central \
  --prop colors=2E75B6,C00000 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop secondaryAxis=2 \
  --prop chartArea.border=D0D0D0:1:solid \
  --prop plotArea.border=E0E0E0:0.5:dot \
  --prop legend=bottom

# ==========================================================================
# Sheet: 6-Advanced
# ==========================================================================
echo "--- 6-Advanced ---"
officecli add "$FILE" / --type sheet --prop name="6-Advanced"

# Chart 1: Reference line with label
# Features: referenceLine (value:color:label:dash style)
officecli add "$FILE" "/6-Advanced" --type chart \
  --prop chartType=bar \
  --prop title="vs Company Average" \
  --prop series1=Score:82,74,91,68,87,72 \
  --prop categories=Engineering,Marketing,Sales,Support,Finance,HR \
  --prop colors=4472C4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop referenceLine=79:FF0000:Average:dash \
  --prop gapwidth=80 \
  --prop gridlines=E0E0E0:0.5:solid

# Chart 2: Conditional coloring (colorRule)
# Features: colorRule (threshold:belowColor:aboveColor), referenceLine=0 (zero baseline)
officecli add "$FILE" "/6-Advanced" --type chart \
  --prop chartType=bar \
  --prop title="Profit/Loss by Division" \
  --prop "series1=P&L:120,85,-45,160,-80,95,-20,140" \
  --prop categories=Div A,Div B,Div C,Div D,Div E,Div F,Div G,Div H \
  --prop colors=2E75B6 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop colorRule=0:C00000:70AD47 \
  --prop referenceLine=0:888888:1:solid \
  --prop dataLabels=outsideEnd \
  --prop labelFont=9:333333:false

# Chart 3: Title glow, title shadow, series shadow
# Features: title.glow (color-radius-opacity), title.shadow, series.shadow on bar charts
officecli add "$FILE" "/6-Advanced" --type chart \
  --prop chartType=bar \
  --prop title="Glow & Shadow Effects" \
  --prop series1=East:185,195,210,228 \
  --prop series2=West:140,152,165,178 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop colors=4472C4,ED7D31 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop title.glow=4472C4-8-60 \
  --prop title.shadow=000000-3-315-2-40 \
  --prop title.font=Calibri --prop title.size=16 \
  --prop title.bold=true --prop title.color=1F4E79 \
  --prop series.shadow=000000-3-315-1-30 \
  --prop plotFill=F0F4F8 --prop chartFill=FFFFFF \
  --prop legend=bottom

# Chart 4: Error bars and data table
# Features: errBars=percent:10, dataTable=true, legend=none
officecli add "$FILE" "/6-Advanced" --type chart \
  --prop chartType=bar \
  --prop title="With Error Bars & Data Table" \
  --prop dataRange=Sheet1!A1:E9 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
  --prop errBars=percent:10 \
  --prop dataTable=true \
  --prop legend=none \
  --prop plotFill=FAFAFA

# ==========================================================================
# Sheet: 7-Axis Controls
# ==========================================================================
echo "--- 7-Axis Controls ---"
officecli add "$FILE" / --type sheet --prop name="7-Axis Controls"

# Chart 1: crosses, crossBetween, valAxisVisible
# Features: crosses=autoZero, crossBetween=between, valAxisVisible=true/false
officecli add "$FILE" "/7-Axis Controls" --type chart \
  --prop chartType=bar \
  --prop title="Axis Cross Controls" \
  --prop series1=Sales:120,80,-30,150 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop crosses=autoZero \
  --prop crossBetween=between \
  --prop valAxisVisible=true

# Chart 2: labelrotation, labeloffset, ticklabelskip
# Features: labelrotation=45, labeloffset=100, ticklabelskip=2
officecli add "$FILE" "/7-Axis Controls" --type chart \
  --prop chartType=column \
  --prop title="Tick-label Rotation, Offset & Skip" \
  --prop series1=Units:45,30,20,55,40,25,60 \
  --prop categories=January,February,March,April,May,June,July \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop labelrotation=45 \
  --prop labeloffset=100 \
  --prop ticklabelskip=2

# Chart 3: axisposition, serlines (stacked bar)
# Features: axisposition=nextTo (alias for tickLabelPos), serlines=true
officecli add "$FILE" "/7-Axis Controls" --type chart \
  --prop chartType=barStacked \
  --prop title="Stacked — axisposition + serlines" \
  --prop series1=Online:55,48,60,70 \
  --prop series2=Retail:30,40,35,25 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop colors=4472C4,ED7D31 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop axisposition=nextTo \
  --prop serlines=true

# Chart 4: markercolor on line/scatter (chart-level fanout)
# Features: markercolor=FF0000 (chart-level fan-out to every series marker)
officecli add "$FILE" "/7-Axis Controls" --type chart \
  --prop chartType=line \
  --prop title="Line — markercolor" \
  --prop series1=Sales:120,145,132,160 \
  --prop series2=Costs:80,95,88,110 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop colors=4472C4,ED7D31 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop marker=circle --prop markerSize=8 \
  --prop markercolor=FF0000 \
  --prop lineWidth=2

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
