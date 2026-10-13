#!/bin/bash
# Doughnut Charts Showcase — generates charts-doughnut.pptx exercising the pptx
# `chart` element with chartType=doughnut across 8 slides.
#
#   Slide 1  holeSize variants      holeSize=10/30/55/75
#   Slide 2  Multi-ring             two-series + three-series concentric rings
#   Slide 3  firstSliceAngle        0 / 90 / 180 / 270
#   Slide 4  Data labels            percent / category / value, leaderlines, labelfont
#   Slide 5  Series styling         colors, gradient, seriesoutline, seriesshadow, transparency
#   Slide 6  Title & legend         title.* + legend positions + legendFont
#   Slide 7  Backgrounds            chartareafill, plotFill, chartborder, roundedcorners
#   Slide 8  Presets & per-series   preset bundles + chart-series Set
#
# CLI twin of charts-doughnut.py (officecli Python SDK).
# Usage: ./charts-doughnut.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/charts-doughnut.pptx"

# shared data
CATS="North,South,East,West"
D="Share:30,25,28,17"
D2="Last:25,30,25,20;This:30,25,28,17"
D3="Region1:30,25,28,17;Region2:25,30,20,25;Region3:20,25,30,25"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ==================== Slide 1: holeSize — 10 / 30 / 55 / 75 ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[1] --type shape --prop text="holeSize — 10 / 30 / 55 / 75" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=holeSize=10 --prop holeSize=10 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=holeSize=30 --prop holeSize=30 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=holeSize=55 --prop holeSize=55 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=holeSize=75 --prop holeSize=75 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"

# ==================== Slide 2: Multi-ring — concentric series ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[2] --type shape --prop text="Multi-ring — concentric series" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="single ring" --prop holeSize=50 --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="two rings" --prop holeSize=40 --prop legend=right --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="three rings" --prop holeSize=30 --prop legend=right --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="two rings + dataLabels=percent" --prop holeSize=40 --prop dataLabels=percent --prop legend=right --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 3: First slice angle — 0 / 90 / 180 / 270 ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[3] --type shape --prop text="First slice angle — 0 / 90 / 180 / 270" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=firstSliceAngle=0 --prop firstSliceAngle=0 --prop holeSize=50 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=firstSliceAngle=90 --prop firstSliceAngle=90 --prop holeSize=50 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=firstSliceAngle=180 --prop firstSliceAngle=180 --prop holeSize=50 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=firstSliceAngle=270 --prop firstSliceAngle=270 --prop holeSize=50 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"

# ==================== Slide 4: Data labels — percent / category / value, leaderlines, labelfont ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[4] --type shape --prop text="Data labels — percent / category / value, leaderlines, labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=dataLabels=percent --prop dataLabels=percent --prop holeSize=50 --prop legend=right --prop labelfont=10:333333:Calibri --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="percent,category" --prop dataLabels="percent,category" --prop holeSize=50 --prop leaderlines=true --prop legend=none --prop labelfont=10:333333:Calibri --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="all flags" --prop dataLabels="value,percent,category" --prop holeSize=50 --prop leaderlines=true --prop legend=none --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=dataLabels=none --prop dataLabels=none --prop holeSize=50 --prop legend=right --prop categories="$CATS" --prop data="$D"

# ==================== Slide 5: Series styling — colors, gradient, outline, shadow, transparency ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[5] --type shape --prop text="Series styling — colors, gradient, outline, shadow, transparency" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=colors= --prop holeSize=50 --prop legend=right --prop colors=4472C4,ED7D31,A5A5A5,70AD47 --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="gradient + seriesshadow" --prop holeSize=50 --prop gradient=FF6600-FFCC00 --prop seriesshadow=000000-5-45-3-50 --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="seriesoutline white" --prop holeSize=50 --prop seriesoutline=FFFFFF:2 --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=transparency=30 --prop holeSize=50 --prop transparency=30 --prop legend=right --prop categories="$CATS" --prop data="$D"

# ==================== Slide 6: Title & legend ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[6] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop holeSize=50 --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="legend=bottom + legendFont" --prop holeSize=50 --prop legend=bottom --prop legendFont=10:333333:Calibri --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="legend.overlay=true" --prop holeSize=50 --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop autotitledeleted=true --prop holeSize=50 --prop legend=none --prop categories="$CATS" --prop data="$D"

# ==================== Slide 7: Backgrounds — chartareafill, plotFill, chartborder, roundedcorners ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[7] --type shape --prop text="Backgrounds — chartareafill, plotFill, chartborder, roundedcorners" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="chartareafill + chartborder" --prop holeSize=50 --prop chartareafill=FFF8E7 --prop chartborder=000000:1 --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=roundedcorners=true --prop holeSize=50 --prop roundedcorners=true --prop chartborder=4472C4:2 --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=plotFill=none --prop holeSize=50 --prop plotFill=none --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title=chartareafill=none --prop holeSize=50 --prop chartareafill=none --prop legend=right --prop categories="$CATS" --prop data="$D"

# ==================== Slide 8: Presets & per-series Set ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[8] --type shape --prop text="Presets & per-series Set" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop preset=minimal --prop title=preset=minimal --prop holeSize=50 --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop preset=dark --prop title=preset=dark --prop holeSize=50 --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop preset=corporate --prop title=preset=corporate --prop holeSize=50 --prop legend=right --prop categories="$CATS" --prop data="$D"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=doughnut --prop title="chart-series Set name+color" --prop holeSize=50 --prop legend=right --prop categories="$CATS" --prop data="$D"

# per-series Set: rename + recolor the single series of the 4th chart
$CLI set "$FILE" /slide[8]/chart[4]/series[1] --prop name="Renamed Share" --prop color=C00000

$CLI close "$FILE"

$CLI validate "$FILE"
echo "Generated: $FILE"
