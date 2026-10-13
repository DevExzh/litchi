#!/bin/bash
# Pie Charts Showcase — pie, pie3d variants across 8 slides.
# CLI twin of charts-pie.py (officecli Python SDK). Both produce an equivalent
# charts-pie.pptx.
#
#   Slide 1  Variants           pie / pie3d (view3d) — varyColors, firstSliceAngle
#   Slide 2  Explosion          explosion=0/10/20/30
#   Slide 3  Title & legend     title.* + legend positions + legendFont
#   Slide 4  Data labels        flags (percent/category/value), labelfont, leaderlines
#   Slide 5  Series styling     colors, gradient, transparency, seriesoutline, seriesshadow
#   Slide 6  First-slice angle  0 / 90 / 180 / 270
#   Slide 7  Backgrounds        chartareafill, plotFill, chartborder, roundedcorners
#   Slide 8  Presets & per-pt   preset bundles + per-point recolor via chart-series Set
#
# Usage:
#   ./charts-pie.sh

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-pie.pptx"
CATS="North,South,East,West"
D="Share:30,25,28,17"

rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"

# ==================== Slide 1: pie variants ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[1] --type shape --prop text="Pie variants — pie / pie3d (varyColors, firstSliceAngle)" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title=pie --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie3d --prop title="pie3d (view3d=20,20,30)" --prop view3d=20,20,30 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="firstSliceAngle=90" --prop firstSliceAngle=90 --prop legend=right --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="varyColors=false" --prop varyColors=false --prop legend=right --prop categories="$CATS" --prop data="$D"

# ==================== Slide 2: explosion ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[2] --type shape --prop text="Explosion — 0 / 10 / 20 / 30 (% of radius)" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="explosion=0" --prop explosion=0 --prop legend=right --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="explosion=10" --prop explosion=10 --prop legend=right --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="explosion=20" --prop explosion=20 --prop legend=right --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="explosion=30" --prop explosion=30 --prop legend=right --prop categories="$CATS" --prop data="$D"

# ==================== Slide 3: title & legend ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[3] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=right --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="legend=bottom + legendFont" --prop legend=bottom --prop legendFont=10:333333:Calibri --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop autotitledeleted=true --prop legend=none --prop categories="$CATS" --prop data="$D"

# ==================== Slide 4: data labels ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[4] --type shape --prop text="Data labels — percent / category / value, labelfont, leaderlines" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="dataLabels=percent" --prop dataLabels=percent --prop legend=right --prop labelfont=10:333333:Calibri --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="percent,category" --prop dataLabels=percent,category --prop leaderlines=true --prop legend=none --prop labelfont=10:333333:Calibri --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="all flags" --prop dataLabels=value,percent,category --prop leaderlines=true --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="dataLabels=none" --prop dataLabels=none --prop legend=right --prop categories="$CATS" --prop data="$D"

# ==================== Slide 5: series styling ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[5] --type shape --prop text="Series styling — colors, gradient, transparency, outline, shadow" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="colors= explicit palette" --prop legend=right --prop colors=4472C4,ED7D31,A5A5A5,70AD47 --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="gradient + seriesshadow" --prop legend=right --prop gradient=FF6600-FFCC00 --prop seriesshadow=000000-5-45-3-50 --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="seriesoutline white" --prop legend=right --prop seriesoutline=FFFFFF:2 --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="transparency=30" --prop legend=right --prop transparency=30 --prop categories="$CATS" --prop data="$D"

# ==================== Slide 6: first-slice angle ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[6] --type shape --prop text="First slice angle — 0 / 90 / 180 / 270" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="firstSliceAngle=0" --prop firstSliceAngle=0 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="firstSliceAngle=90" --prop firstSliceAngle=90 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="firstSliceAngle=180" --prop firstSliceAngle=180 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="firstSliceAngle=270" --prop firstSliceAngle=270 --prop legend=right --prop varyColors=true --prop categories="$CATS" --prop data="$D"

# ==================== Slide 7: backgrounds ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[7] --type shape --prop text="Backgrounds — chartareafill, plotFill, chartborder, roundedcorners" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="chartareafill + chartborder" --prop legend=right --prop chartareafill=FFF8E7 --prop chartborder=000000:1 --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="roundedcorners=true" --prop legend=right --prop roundedcorners=true --prop chartborder=4472C4:2 --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="plotFill=none" --prop legend=right --prop plotFill=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="chartareafill=none" --prop legend=right --prop chartareafill=none --prop categories="$CATS" --prop data="$D"

# ==================== Slide 8: presets & per-series Set ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[8] --type shape --prop text="Presets & per-series Set" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop preset=minimal --prop title="preset=minimal" --prop legend=right --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=pie --prop preset=dark --prop title="preset=dark" --prop legend=right --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop preset=corporate --prop title="preset=corporate" --prop legend=right --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie --prop title="chart-series Set name+color" --prop legend=right --prop categories="$CATS" --prop data="$D"
# per-point recolor via chart-series Set (must follow the chart[4] add above)
officecli set "$FILE" /slide[8]/chart[4]/series[1] --prop name="Renamed Share" --prop color=C00000

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
