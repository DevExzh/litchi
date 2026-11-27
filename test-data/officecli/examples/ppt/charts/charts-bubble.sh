#!/bin/bash
# Bubble Charts Showcase — generates charts-bubble.pptx exercising the pptx
# `chart` element with chartType=bubble across the full styling surface.
#
# CLI twin of charts-bubble.py (officecli Python SDK). Both produce an
# equivalent charts-bubble.pptx.
#
#   Slide 1  bubbleScale            50 / 100 / 150 / 200 (% of default)
#   Slide 2  sizerepresents         area vs width
#   Slide 3  shownegbubbles         true vs false (with negative values)
#   Slide 4  Title & legend         title.* + legend positions + legendFont
#   Slide 5  Data labels            value/category/bubbleSize, labelfont
#   Slide 6  Axes                   min/max, gridlines, ticks
#   Slide 7  Series styling         colors, gradient, transparency, outline, shadow
#   Slide 8  Presets & per-series   preset bundles + chart-series Set
#
# Usage: ./charts-bubble.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/charts-bubble.pptx"

# Quadrant boxes (reused on every slide).
TL=(--prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in)
TR=(--prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in)
BL=(--prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in)
BR=(--prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in)
D="A:5,12,8,18,22,9,15,11"
D2="A:5,12,8,18,22,9;B:7,11,15,9,20,14"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ==================== Slide 1: bubbleScale ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[1] --type shape --prop text="bubbleScale — 50 / 100 / 150 / 200 (% of default)" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[1] --type chart "${TL[@]}" --prop chartType=bubble --prop title=bubbleScale=50  --prop bubbleScale=50  --prop legend=none --prop data="$D"
$CLI add "$FILE" /slide[1] --type chart "${TR[@]}" --prop chartType=bubble --prop title=bubbleScale=100 --prop bubbleScale=100 --prop legend=none --prop data="$D"
$CLI add "$FILE" /slide[1] --type chart "${BL[@]}" --prop chartType=bubble --prop title=bubbleScale=150 --prop bubbleScale=150 --prop legend=none --prop data="$D"
$CLI add "$FILE" /slide[1] --type chart "${BR[@]}" --prop chartType=bubble --prop title=bubbleScale=200 --prop bubbleScale=200 --prop legend=none --prop data="$D"

# ==================== Slide 2: sizerepresents ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[2] --type shape --prop text="sizerepresents — area vs width" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[2] --type chart "${TL[@]}" --prop chartType=bubble --prop title="sizerepresents=area"  --prop sizerepresents=area  --prop legend=none   --prop data="$D"
$CLI add "$FILE" /slide[2] --type chart "${TR[@]}" --prop chartType=bubble --prop title="sizerepresents=width" --prop sizerepresents=width --prop legend=none   --prop data="$D"
$CLI add "$FILE" /slide[2] --type chart "${BL[@]}" --prop chartType=bubble --prop title="area + 2 series"      --prop sizerepresents=area  --prop legend=bottom --prop data="$D2"
$CLI add "$FILE" /slide[2] --type chart "${BR[@]}" --prop chartType=bubble --prop title="width + 2 series"     --prop sizerepresents=width --prop legend=bottom --prop data="$D2"

# ==================== Slide 3: shownegbubbles ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[3] --type shape --prop text="shownegbubbles — false vs true" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[3] --type chart "${TL[@]}" --prop chartType=bubble --prop title="shownegbubbles=false" --prop shownegbubbles=false --prop legend=none   --prop data="A:5,-8,12,-15,18,22"
$CLI add "$FILE" /slide[3] --type chart "${TR[@]}" --prop chartType=bubble --prop title="shownegbubbles=true"  --prop shownegbubbles=true  --prop legend=none   --prop data="A:5,-8,12,-15,18,22"
$CLI add "$FILE" /slide[3] --type chart "${BL[@]}" --prop chartType=bubble --prop title="false + 2 series"     --prop shownegbubbles=false --prop legend=bottom --prop data="A:5,-8,12,-15,18,22;B:8,11,-9,14,-16,20"
$CLI add "$FILE" /slide[3] --type chart "${BR[@]}" --prop chartType=bubble --prop title="true + 2 series"      --prop shownegbubbles=true  --prop legend=bottom --prop data="A:5,-8,12,-15,18,22;B:8,11,-9,14,-16,20"

# ==================== Slide 4: Title & legend ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[4] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[4] --type chart "${TL[@]}" --prop chartType=bubble --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart "${TR[@]}" --prop chartType=bubble --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart "${BL[@]}" --prop chartType=bubble --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart "${BR[@]}" --prop chartType=bubble --prop autotitledeleted=true --prop legend=none --prop data="$D2"

# ==================== Slide 5: Data labels ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[5] --type shape --prop text="Data labels — flags + labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[5] --type chart "${TL[@]}" --prop chartType=bubble --prop title=value --prop dataLabels=value --prop labelfont=9:333333:Calibri --prop legend=none --prop data="$D"
$CLI add "$FILE" /slide[5] --type chart "${TR[@]}" --prop chartType=bubble --prop title="value,series" --prop dataLabels=value,series --prop legend=none --prop data="$D2"
$CLI add "$FILE" /slide[5] --type chart "${BL[@]}" --prop chartType=bubble --prop title="labelPos=top" --prop dataLabels=value --prop labelPos=top --prop legend=none --prop data="$D"
$CLI add "$FILE" /slide[5] --type chart "${BR[@]}" --prop chartType=bubble --prop title="dataLabels=none" --prop dataLabels=none --prop legend=none --prop data="$D"

# ==================== Slide 6: Axes ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[6] --type shape --prop text="Axes — min/max, gridlines, ticks" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[6] --type chart "${TL[@]}" --prop chartType=bubble --prop title="min/max + titles" --prop axismin=0 --prop axismax=30 --prop majorunit=10 --prop axistitle=Y --prop cattitle=X --prop axisfont=10:333333:Calibri --prop axisline=666666:1 --prop legend=none --prop data="$D"
$CLI add "$FILE" /slide[6] --type chart "${TR[@]}" --prop chartType=bubble --prop title="gridlines + minorGridlines" --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop legend=none --prop data="$D"
$CLI add "$FILE" /slide[6] --type chart "${BL[@]}" --prop chartType=bubble --prop title="labelrotation=-30" --prop labelrotation=-30 --prop legend=none --prop data="$D"
$CLI add "$FILE" /slide[6] --type chart "${BR[@]}" --prop chartType=bubble --prop title="dispunits=hundreds" --prop dispunits=hundreds --prop legend=none --prop data="A:500,1200,800,1800,2200,900"

# ==================== Slide 7: Series styling ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[7] --type shape --prop text="Series styling — colors, gradient, transparency, outline, shadow" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[7] --type chart "${TL[@]}" --prop chartType=bubble --prop title="colors + seriesoutline" --prop colors=4472C4,ED7D31 --prop seriesoutline=000000:0.5 --prop legend=bottom --prop data="$D2"
$CLI add "$FILE" /slide[7] --type chart "${TR[@]}" --prop chartType=bubble --prop title="gradient + seriesshadow" --prop gradient=FF6600-FFCC00 --prop seriesshadow=000000-5-45-3-50 --prop legend=none --prop data="$D"
$CLI add "$FILE" /slide[7] --type chart "${BL[@]}" --prop chartType=bubble --prop title="transparency=30" --prop transparency=30 --prop legend=bottom --prop data="$D2"
$CLI add "$FILE" /slide[7] --type chart "${BR[@]}" --prop chartType=bubble --prop title="per-series gradients" --prop gradients="FF0000-0000FF;00FF00-FFFF00" --prop legend=bottom --prop data="$D2"

# ==================== Slide 8: Presets & per-series Set ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[8] --type shape --prop text="Presets & per-series Set" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[8] --type chart "${TL[@]}" --prop chartType=bubble --prop preset=minimal   --prop title="preset=minimal"   --prop legend=bottom --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart "${TR[@]}" --prop chartType=bubble --prop preset=dark      --prop title="preset=dark"      --prop legend=bottom --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart "${BL[@]}" --prop chartType=bubble --prop preset=corporate --prop title="preset=corporate" --prop legend=bottom --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart "${BR[@]}" --prop chartType=bubble --prop title="chart-series Set name+color" --prop legend=bottom --prop data="$D2"

# chart-series Set (slide 8, chart[4]) — runs after the chart exists.
$CLI set "$FILE" /slide[8]/chart[4]/series[1] --prop name="Renamed A" --prop color=C00000
$CLI set "$FILE" /slide[8]/chart[4]/series[2] --prop name="Renamed B" --prop color=2E75B6

$CLI close "$FILE"
$CLI validate "$FILE"
echo "Generated: $FILE"
