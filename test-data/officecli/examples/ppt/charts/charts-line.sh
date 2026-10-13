#!/bin/bash
# Line Charts Showcase — line / stackedLine / percentStackedLine / line3d.
# CLI twin of charts-line.py (officecli Python SDK). Both produce an equivalent
# charts-line.pptx.
#
#   Slide 1  Variants     line / stackedLine / percentStackedLine / line3d
#   Slide 2  Markers      marker symbol/size/color, markersize, showMarker
#   Slide 3  Smoothing    smooth, linedash, linewidth
#   Slide 4  Title&legend title.* + legend positions + legendFont
#   Slide 5  Data labels  flags, labelPos, labelfont
#   Slide 6  Axes         min/max, titles, fonts, gridlines, ticks, labelrotation, log
#   Slide 7  Overlays     droplines, hilowlines, updownbars, trendline, errbars, referenceline
#   Slide 8  Per-series   lineWidth/lineDash/marker/markerSize/color/smooth + presets
#
# Usage: ./charts-line.sh

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/charts-line.pptx"

CATS="Mon,Tue,Wed,Thu,Fri"
D2="A:50,60,70,65,80;B:40,45,55,60,75"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# title shape helper is inlined per slide (bash has no kwargs).

# ==================== Slide 1: Line variants ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[1] --type shape --prop text="Line variants — line / stackedLine / percentStackedLine / line3d" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=line --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stackedLine --prop title=stackedLine --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=percentStackedLine --prop title=percentStackedLine --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line3d --prop title=line3d --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 2: Markers ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[2] --type shape --prop text="Markers — symbol, size, color, showMarker" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=marker=circle:8:FF0000 --prop marker=circle:8:FF0000 --prop linewidth=2 --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=marker=square:6 --prop marker=square:6 --prop linewidth=2 --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=marker=diamond:10:0070C0 --prop marker=diamond:10:0070C0 --prop linewidth=2 --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="showMarker=true (default markers)" --prop showMarker=true --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 3: Smoothing & dash ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[3] --type shape --prop text="Smoothing & dash — smooth, linedash, linewidth" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=smooth=true --prop smooth=true --prop linewidth=2.5 --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=linedash=dash --prop linedash=dash --prop linewidth=2 --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=linedash=dot --prop linedash=dot --prop linewidth=2 --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=linedash=dashDot --prop linedash=dashDot --prop linewidth=2 --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"

# ==================== Slide 4: Title & legend ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[4] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop autotitledeleted=true --prop legend=none --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 5: Data labels ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[5] --type shape --prop text="Data labels — flags, labelPos, labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="dataLabels=value @ top" --prop dataLabels=value --prop labelPos=top --prop labelfont=10:333333:Calibri --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=value,category --prop dataLabels=value,category --prop labelPos=top --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=dataLabels=none --prop dataLabels=none --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="labelfont styled" --prop dataLabels=value --prop labelPos=top --prop labelfont=12:C00000:Georgia --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"

# ==================== Slide 6: Axes ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[6] --type shape --prop text="Axes — min/max, gridlines, ticks, labelrotation, log" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="min/max + titles" --prop legend=none --prop axismin=0 --prop axismax=100 --prop majorunit=25 --prop axistitle=Visits --prop cattitle=Day --prop axisfont=10:333333:Calibri --prop axisline=666666:1 --prop axisnumfmt="#,##0" --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="gridlines + ticks" --prop legend=none --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop majorTickMark=out --prop minorTickMark=in --prop tickLabelPos=nextTo --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=labelrotation=-30 --prop legend=none --prop labelrotation=-30 --prop categories="January,February,March,April,May,June" --prop data="A:60,90,140,180,160,210"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=logbase=10 --prop legend=none --prop logbase=10 --prop axismin=1 --prop axismax=10000 --prop categories="$CATS" --prop data="Growth:5,50,500,5000,3000"

# ==================== Slide 7: Overlays ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[7] --type shape --prop text="Overlays — droplines, hilowlines, updownbars, trendline, errbars, referenceline" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="droplines + hilowlines" --prop droplines=808080:0.5 --prop hilowlines=true --prop legend=bottom --prop categories="$CATS" --prop data="High:130,135,140,138,145;Low:118,122,128,125,132"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=updownbars=150:00AA00:FF0000 --prop updownbars=150:00AA00:FF0000 --prop legend=bottom --prop categories="$CATS" --prop data="Open:120,128,130,135,138;Close:128,125,135,138,142"
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="trendline=linear + errbars=stdDev:1" --prop trendline=linear --prop errbars=stdDev:1 --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title=referenceline=70:FF0000:Target --prop referenceline=70:FF0000:Target --prop legend=none --prop categories="$CATS" --prop data="A:50,60,70,65,80"

# ==================== Slide 8: Per-series Set + presets ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[8] --type shape --prop text="Per-series Set + presets — chart-series lineWidth/lineDash/marker/markerSize/color/smooth" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop preset=minimal --prop title=preset=minimal --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=line --prop preset=dark --prop title=preset=dark --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop preset=corporate --prop title=preset=corporate --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line --prop title="chart-series Set per line" --prop showMarker=true --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# chart-series Set on the 4th chart's two lines (after the chart exists).
$CLI set "$FILE" /slide[8]/chart[4]/series[1] --prop name=Alpha --prop color=C00000 --prop lineWidth=2.5 --prop lineDash=solid --prop marker=circle --prop markerSize=9 --prop smooth=true
$CLI set "$FILE" /slide[8]/chart[4]/series[2] --prop name=Beta --prop color=2E75B6 --prop lineWidth=1.5 --prop lineDash=dash --prop marker=diamond --prop markerSize=8

$CLI close "$FILE"
$CLI validate "$FILE"
echo "Generated: $FILE"
