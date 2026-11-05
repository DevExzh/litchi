#!/bin/bash
# Combo Charts Showcase — combotypes, combosplit, secondaryaxis.
# Generates: charts-combo.pptx
#
#   Slide 1  combotypes mixes       column+line, column+area, line+area, bar+line
#   Slide 2  combosplit             split index 1, 2, 3 (first N series use primary)
#   Slide 3  secondaryaxis          1 series, 2 series, multiple series on secondary
#   Slide 4  Title & legend
#   Slide 5  Data labels
#   Slide 6  Axes                   min/max on both axes, titles, gridlines
#   Slide 7  Series styling         colors, gradients, transparency, outline, shadow
#   Slide 8  Presets & per-series   preset bundles + chart-series Set
#
# CLI twin of charts-combo.py (officecli Python SDK). Both produce an
# equivalent charts-combo.pptx.
#
# Usage:
#   ./charts-combo.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/charts-combo.pptx"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# Shared category + data strings (reused across every chart).
CATS="Q1,Q2,Q3,Q4"
D2="Sales:120,135,148,162;Growth %:5,12,18,22"
D3="Sales:120,135,148,162;Cost:80,90,95,105;Growth %:5,12,18,22"

# ==================== Slide 1: combotypes ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[1] --type shape --prop text="combotypes — column+line / column+area / line+area / bar+line" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="column + line" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,area --prop title="column + area" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=line,area --prop title="line + area" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=bar,line --prop title="bar + line" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 2: combosplit ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[2] --type shape --prop text="combosplit — first N series use primary type" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,column,line --prop combosplit=2 --prop title="combosplit=2 (2 columns + 1 line)" --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line,line --prop combosplit=1 --prop title="combosplit=1 (1 column + 2 lines)" --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=line,line,column --prop combosplit=2 --prop title="combosplit=2 (2 lines + 1 column)" --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=area,column,line --prop combosplit=1 --prop title="area + column + line" --prop legend=bottom --prop categories="$CATS" --prop data="$D3"

# ==================== Slide 3: secondaryaxis ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[3] --type shape --prop text="secondaryaxis — line on secondary value axis" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop secondaryaxis=2 --prop title="secondaryaxis=2" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,column,line --prop secondaryaxis=3 --prop combosplit=2 --prop title="secondaryaxis=3 (Growth on right)" --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line,line --prop secondaryaxis=2,3 --prop combosplit=1 --prop title="secondaryaxis=2,3" --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop secondaryaxis=2 --prop title="with grid + tick fonts" --prop gridlines=E0E0E0:0.3 --prop axisfont=9:333333:Calibri --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 4: Title & legend ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[4] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop autotitledeleted=true --prop legend=none --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 5: Data labels ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[5] --type shape --prop text="Data labels — combo charts skip labelPos (chart-type conditional)" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="dataLabels=value" --prop dataLabels=value --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="value,series" --prop dataLabels=value,series --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="dataLabels=none" --prop dataLabels=none --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="labelfont styled" --prop dataLabels=value --prop labelfont=10:C00000:Georgia --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 6: Axes ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[6] --type shape --prop text="Axes — min/max on primary, secondary, gridlines, axisnumfmt" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop secondaryaxis=2 --prop title="both axes min/max" --prop axismin=0 --prop axismax=200 --prop axistitle=Sales --prop cattitle=Quarter --prop axisfont=10:333333:Calibri --prop axisnumfmt="#,##0" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="gridlines + minorGridlines" --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="labelrotation=-30" --prop labelrotation=-30 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="chart-axis Set after add" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI set "$FILE" "/slide[6]/chart[4]/axis[@role=value]" --prop title="Sales (USD)" --prop format="\$#,##0" --prop majorGridlines=true --prop min=0 --prop max=200

# ==================== Slide 7: Series styling ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[7] --type shape --prop text="Series styling — colors, gradient(s), transparency, outline, shadow" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="colors + seriesoutline" --prop colors=4472C4,ED7D31 --prop seriesoutline=000000:0.5 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="gradient + seriesshadow" --prop gradient=FF6600-FFCC00 --prop seriesshadow=000000-5-45-3-50 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="transparency=30" --prop transparency=30 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="per-series gradients" --prop gradients="FF0000-0000FF;00FF00-FFFF00" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 8: Presets & per-series Set ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[8] --type shape --prop text="Presets & per-series Set" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop preset=minimal --prop title="preset=minimal" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop preset=dark --prop title="preset=dark" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop preset=corporate --prop title="preset=corporate" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop title="chart-series Set" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI set "$FILE" /slide[8]/chart[4]/series[1] --prop name="Renamed Sales" --prop color=C00000
$CLI set "$FILE" /slide[8]/chart[4]/series[2] --prop name="Renamed Growth" --prop color=2E75B6 --prop lineWidth=2.5 --prop marker=circle --prop markerSize=8

$CLI close "$FILE"

$CLI validate "$FILE"
echo "Generated: $FILE"
