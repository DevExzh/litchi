#!/bin/bash
# Area Charts Showcase — generates charts-area.pptx exercising the pptx `chart`
# element across the area family (area / stackedArea / percentStackedArea /
# area3d) plus titles, legends, data labels, axes, series styling, overlays,
# backgrounds, presets and per-series Set.
#
# CLI twin of charts-area.py (officecli Python SDK). Both produce an equivalent
# charts-area.pptx — this one issues one `officecli` process per command.
#
#   Slide 1  Variants       area / stackedArea / percentStackedArea / area3d
#   Slide 2  Title & legend
#   Slide 3  Data labels
#   Slide 4  Axes
#   Slide 5  Series styling
#   Slide 6  Overlays
#   Slide 7  Backgrounds
#   Slide 8  Presets & per-series control
#
# Usage: ./charts-area.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/charts-area.pptx"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ==================== Slide 1: variants ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[1] --type shape --prop text="Area variants — area / stackedArea / percentStackedArea / area3d" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title=area --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stackedArea --prop title=stackedArea --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=percentStackedArea --prop title=percentStackedArea --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area3d --prop title=area3d --prop view3d=15,20,30 --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'

# ==================== Slide 2: title & legend ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[2] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop autotitledeleted=true --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'

# ==================== Slide 3: data labels ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[3] --type shape --prop text="Data labels — flags, labelPos, labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="dataLabels=value" --prop dataLabels=value --prop labelfont=10:333333:Calibri --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stackedArea --prop title="stacked + center labels" --prop dataLabels=value --prop labelPos=center --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="value,category" --prop dataLabels=value,category --prop labelfont=9:333333:Calibri --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="dataLabels=none" --prop dataLabels=none --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80

# ==================== Slide 4: axes ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[4] --type shape --prop text="Axes — min/max, gridlines, ticks, labelrotation" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="min/max + titles" --prop legend=none --prop axismin=0 --prop axismax=100 --prop majorunit=25 --prop axistitle=Value --prop cattitle=Day --prop axisfont=10:333333:Calibri --prop axisline=666666:1 --prop 'axisnumfmt=#,##0' --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="gridlines + ticks" --prop legend=none --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop majorTickMark=out --prop minorTickMark=in --prop tickLabelPos=nextTo --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="labelrotation=-30" --prop legend=none --prop labelrotation=-30 --prop categories=January,February,March,April,May,June --prop data=A:60,90,140,180,160,210
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="dispunits=thousands" --prop legend=none --prop dispunits=thousands --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=Rev:120000,135000,148000,162000,180000

# ==================== Slide 5: series styling ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[5] --type shape --prop text="Series styling — colors, gradient(s), transparency, outline, shadow" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="colors + seriesoutline" --prop legend=bottom --prop colors=4472C4,ED7D31 --prop seriesoutline=000000:0.5 --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="gradient + seriesshadow" --prop legend=none --prop gradient=FF6600-FFCC00:90 --prop seriesshadow=000000-5-45-3-50 --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="per-series gradients + transparency=30" --prop 'gradients=FF0000-0000FF;00FF00-FFFF00' --prop transparency=30 --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="single + transparency=50" --prop transparency=50 --prop colors=4472C4 --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80

# ==================== Slide 6: overlays ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[6] --type shape --prop text="Overlays — referenceline, errbars, trendline" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="referenceline=60" --prop referenceline=60:FF0000:Target --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="errbars=percentage:10" --prop errbars=percentage:10 --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="trendline=linear" --prop trendline=linear --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="trendline=movingAvg:3" --prop trendline=movingAvg:3 --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80

# ==================== Slide 7: backgrounds ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[7] --type shape --prop text="Backgrounds — chartareafill, plotFill, chartborder, plotborder, roundedcorners" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="chartareafill + plotFill + borders" --prop legend=bottom --prop chartareafill=FFF8E7 --prop plotFill=FAFAFA --prop chartborder=000000:1 --prop plotborder=CCCCCC:0.5 --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="roundedcorners=true" --prop roundedcorners=true --prop chartborder=4472C4:2 --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="plotFill=none" --prop plotFill=none --prop gridlines=none --prop legend=none --prop categories=Mon,Tue,Wed,Thu,Fri --prop data=A:50,60,70,65,80
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="dataTable=true" --prop dataTable=true --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'

# ==================== Slide 8: presets & per-series control ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[8] --type shape --prop text="Presets & per-series control" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop preset=minimal --prop title=preset=minimal --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area --prop preset=dark --prop title=preset=dark --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop preset=corporate --prop title=preset=corporate --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop 'data=Web:50,60,70,65,80;Mobile:30,35,42,48,55'
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=area --prop title="seriesN.* + chart-series Set" --prop legend=bottom --prop categories=Mon,Tue,Wed,Thu,Fri --prop series1.name=Web --prop series1.values=50,60,70,65,80 --prop series1.color=4472C4 --prop series2.name=Mobile --prop series2.values=30,35,42,48,55 --prop series2.color=ED7D31
# chart-series Set — recolour/rename series[1] of the BR chart after Add
$CLI set "$FILE" /slide[8]/chart[4]/series[1] --prop name="Renamed Web" --prop color=C00000

$CLI close "$FILE"

$CLI validate "$FILE"
echo "Generated: $FILE"
