#!/bin/bash
# Scatter Charts Showcase — scatterstyle line/lineMarker/marker/smooth/smoothMarker.
# CLI twin of charts-scatter.py (officecli Python SDK). Both produce an
# equivalent charts-scatter.pptx.
#
#   Slide 1  scatterstyle variants  line / lineMarker / marker / smooth / smoothMarker
#   Slide 2  Markers                marker symbol/size/color/markercolor
#   Slide 3  Title & legend         title.overlay / legend.overlay
#   Slide 4  Data labels            flags + labelfont
#   Slide 5  Axes                   min/max, gridlines, log
#   Slide 6  Series styling         colors, gradient, transparency, outline, shadow
#   Slide 7  Overlays               trendline / errbars
#   Slide 8  Per-series Set         lineWidth/lineDash/marker/markerSize/color/smooth + presets
#   Slide 9  series{N}=             named series shorthand
#
# Usage:
#   ./charts-scatter.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-scatter.pptx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# Shared scatter data
D="A:10,20,18,30,28,40,42,55,52,65"
D2="A:10,20,18,30,28,40,42,55;B:5,12,15,22,25,30,35,40"

# ==================== Slide 1: scatterstyle variants ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[1] --type shape --prop text="scatterstyle — line / lineMarker / marker / smooth / smoothMarker" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[1] --type chart --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=line        --prop title="scatterstyle=line"        --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=lineMarker  --prop title="scatterstyle=lineMarker"  --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[1] --type chart --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker      --prop title="scatterstyle=marker"      --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=smoothMarker --prop title="scatterstyle=smoothMarker" --prop legend=none --prop data="$D"

# ==================== Slide 2: Markers ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[2] --type shape --prop text="Markers — symbol / size / color / markercolor" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[2] --type chart --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="circle:10:FF0000"  --prop marker=circle:10:FF0000  --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="diamond:12:0070C0" --prop marker=diamond:12:0070C0 --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[2] --type chart --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="square:8:70AD47"   --prop marker=square:8:70AD47   --prop legend=none --prop data="$D"
# markercolor — per-series marker fill color (independent of marker= compound form)
officecli add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="markercolor=E63946" --prop marker=circle:10 --prop markercolor=E63946 --prop legend=none --prop data="$D"

# ==================== Slide 3: Title & legend ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[3] --type shape --prop text="Title & legend — title.overlay / legend.overlay" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[3] --type chart --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=smoothMarker --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop data="$D2"
officecli add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=lineMarker --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop data="$D2"
officecli add "$FILE" /slide[3] --type chart --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=lineMarker --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop data="$D2"
# title.overlay — title rendered over the plot area (saves vertical space)
officecli add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="title.overlay=true" --prop title.overlay=true --prop legend=none --prop data="$D2"

# ==================== Slide 4: Data labels ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[4] --type shape --prop text="Data labels — flags + labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[4] --type chart --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="value" --prop dataLabels=value --prop labelfont=9:333333:Calibri --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="value,series" --prop dataLabels=value,series --prop legend=none --prop data="$D2"
officecli add "$FILE" /slide[4] --type chart --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="labelPos=top" --prop dataLabels=value --prop labelPos=top --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="dataLabels=none" --prop dataLabels=none --prop legend=none --prop data="$D"

# ==================== Slide 5: Axes ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[5] --type shape --prop text="Axes — min/max, gridlines, ticks, log on both axes" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[5] --type chart --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=lineMarker --prop title="min/max + titles" --prop axismin=0 --prop axismax=80 --prop majorunit=20 --prop axistitle=Y --prop cattitle=X --prop axisfont=10:333333:Calibri --prop axisline=666666:1 --prop axisnumfmt="#,##0" --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="gridlines + minorGridlines" --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="labelrotation=-30" --prop labelrotation=-30 --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="logbase=10 (Y)" --prop logbase=10 --prop axismin=1 --prop axismax=100 --prop legend=none --prop data="A:2,5,8,12,20,40,80"

# ==================== Slide 6: Series styling ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[6] --type shape --prop text="Series styling — colors, gradient, transparency, outline, shadow" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[6] --type chart --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="colors + seriesoutline" --prop colors=4472C4,ED7D31 --prop seriesoutline=000000:0.5 --prop legend=bottom --prop data="$D2"
officecli add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="gradient + seriesshadow" --prop gradient=FF6600-FFCC00 --prop seriesshadow=000000-5-45-3-50 --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="transparency=30" --prop transparency=30 --prop legend=bottom --prop data="$D2"
officecli add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="per-series gradients" --prop gradients="FF0000-0000FF;00FF00-FFFF00" --prop legend=bottom --prop data="$D2"

# ==================== Slide 7: Overlays ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[7] --type shape --prop text="Overlays — trendline (linear/poly/exp/movingAvg), errbars, referenceline" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[7] --type chart --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="trendline=linear" --prop trendline=linear --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="trendline=poly:3" --prop trendline=poly:3 --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[7] --type chart --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="trendline=movingAvg:3" --prop trendline=movingAvg:3 --prop legend=none --prop data="$D"
officecli add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="errbars=stdDev:1" --prop errbars=stdDev:1 --prop legend=none --prop data="$D"

# ==================== Slide 8: Per-series Set + presets ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[8] --type shape --prop text="Per-series Set + presets — lineWidth/lineDash/marker/markerSize/color/smooth" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[8] --type chart --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=smoothMarker --prop preset=minimal   --prop title="preset=minimal"   --prop legend=bottom --prop data="$D2"
officecli add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=smoothMarker --prop preset=dark      --prop title="preset=dark"      --prop legend=bottom --prop data="$D2"
officecli add "$FILE" /slide[8] --type chart --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=smoothMarker --prop preset=corporate --prop title="preset=corporate" --prop legend=bottom --prop data="$D2"
officecli add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=lineMarker --prop title="chart-series Set per series" --prop legend=bottom --prop data="$D2"
# chart-series Set per series (path-based set, after the chart exists)
officecli set "$FILE" /slide[8]/chart[4]/series[1] --prop name=Alpha --prop color=C00000 --prop lineWidth=2.5 --prop lineDash=solid --prop marker=circle  --prop markerSize=10 --prop smooth=true
officecli set "$FILE" /slide[8]/chart[4]/series[2] --prop name=Beta  --prop color=2E75B6 --prop lineWidth=1.5 --prop lineDash=dash  --prop marker=diamond --prop markerSize=8

# ==================== Slide 9: series{N}= named series shorthand ====================
# series{N}= is an alternative to data= that names each series at Add time.
# series1=Name:v1,v2,…  series2=Name:v1,v2,…  (no shared categories needed for scatter)
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[9] --type shape --prop text="series{N}= — named series shorthand (name:v1,v2,…)" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[9] --type chart --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=lineMarker --prop title="series1= + series2=" --prop series1="Alpha:10,25,18,40" --prop series2="Beta:5,15,12,30" --prop legend=bottom
officecli add "$FILE" /slide[9] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="three named series" --prop series1="Group A:8,20,15" --prop series2="Group B:4,12,10" --prop series3="Group C:12,28,22" --prop legend=bottom
officecli add "$FILE" /slide[9] --type chart --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=smoothMarker --prop title="series1 with colors" --prop series1="Rev:30,45,55,70" --prop series2="Cost:20,30,35,42" --prop colors=4472C4,E63946 --prop legend=bottom
officecli add "$FILE" /slide[9] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=scatter --prop scatterstyle=marker --prop title="series1.* per-series naming + colors=" --prop series1.name=Alpha --prop series1.values="10,25,18,40" --prop series2.name=Beta --prop series2.values="5,15,12,30" --prop colors=4472C4,E63946 --prop legend=bottom

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
