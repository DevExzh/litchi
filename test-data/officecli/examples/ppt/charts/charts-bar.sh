#!/bin/bash
# Bar Charts Showcase — bar, stackedBar, percentStackedBar, bar3d (cylinder/cone/pyramid).
# Generates: charts-bar.pptx
#
#   Slide 1  Variants           bar / stackedBar / percentStackedBar / bar3d
#   Slide 2  3D bar shapes      shape=box/cylinder/cone/pyramid (bar3d only)
#   Slide 3  Title & legend     title.* + legend positions + legendFont
#   Slide 4  Data labels        flags + labelPos + labelfont
#   Slide 5  Axes               min/max/title/font/line/numfmt/gridlines/labelrotation
#   Slide 6  Series styling     colors, gradient, transparency, outline, shadow, invertifneg, serlines
#   Slide 7  Overlays           referenceline, errbars, gapwidth, overlap, dataTable
#   Slide 8  Presets & per-ser  preset bundles + seriesN.* + chart-series Set
#
# CLI twin of charts-bar.py (officecli Python SDK).
# Usage: ./charts-bar.sh [officecli path]

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/charts-bar.pptx"

# shared geometry + data
CATS="Q1,Q2,Q3,Q4"
D2="East:120,135,148,162;West:95,108,115,128"
D3="East:120,135,148,162;South:95,108,115,128;West:80,90,98,110"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ==================== Slide 1: Bar variants ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[1] --type shape --prop text="Bar variants — bar / stackedBar / percentStackedBar / bar3d" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title=bar --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stackedBar --prop title=stackedBar --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=percentStackedBar --prop title=percentStackedBar --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop title=bar3d --prop legend=bottom --prop categories="$CATS" --prop data="$D2" --prop view3d=15,20,30

# ==================== Slide 2: 3D bar shapes ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[2] --type shape --prop text="3D bar shapes — shape=box / cylinder / cone / pyramid" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop shape=box --prop title=shape=box --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop shape=cylinder --prop title=shape=cylinder --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop shape=cone --prop title=shape=cone --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop shape=pyramid --prop title=shape=pyramid --prop legend=none --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 3: Title & legend ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[3] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop autotitledeleted=true --prop legend=none --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 4: Data labels ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[4] --type shape --prop text="Data labels — flags, labelPos, labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="value @ outsideEnd" --prop dataLabels=value --prop labelPos=outsideEnd --prop labelfont=10:333333:Calibri --prop legend=none --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="value,category @ insideEnd" --prop dataLabels=value,category --prop labelPos=insideEnd --prop labelfont=9:FFFFFF:Calibri --prop legend=none --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stackedBar --prop title="stacked + center labels" --prop dataLabels=value --prop labelPos=center --prop labelfont=9:FFFFFF:Calibri --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="dataLabels=none" --prop dataLabels=none --prop legend=none --prop categories="$CATS" --prop data="A:60,90,140,180"

# ==================== Slide 5: Axes ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[5] --type shape --prop text="Axes — min/max, titles, fonts, gridlines, ticks, labelrotation" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="min/max + titles + numfmt" --prop legend=none --prop axismin=0 --prop axismax=200 --prop majorunit=50 --prop minorunit=10 --prop axistitle=Revenue --prop cattitle=Quarter --prop axisfont=10:333333:Calibri --prop axisline=666666:1 --prop axisnumfmt="#,##0" --prop categories="$CATS" --prop data="Rev:60,90,140,180"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="gridlines + ticks" --prop legend=none --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop majorTickMark=out --prop minorTickMark=in --prop tickLabelPos=nextTo --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="labelrotation=-30" --prop legend=none --prop labelrotation=-30 --prop categories="January,February,March,April" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="dispunits=thousands" --prop legend=none --prop dispunits=thousands --prop categories="$CATS" --prop data="Rev:120000,135000,148000,162000"
$CLI set "$FILE" "/slide[5]/chart[1]/axis[@role=value]" --prop title=Revenue --prop format="\$#,##0" --prop majorGridlines=true --prop max=200 --prop min=0

# ==================== Slide 6: Series styling ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[6] --type shape --prop text="Series styling — colors, gradient(s), transparency, outline, shadow, invertifneg, serlines" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="colors + seriesoutline" --prop legend=bottom --prop colors=4472C4,ED7D31,A5A5A5 --prop seriesoutline=000000:0.5 --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="gradient + seriesshadow" --prop legend=bottom --prop gradient=FF6600-FFCC00:90 --prop seriesshadow=000000-5-45-3-50 --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="transparency=30 + gradients" --prop legend=bottom --prop gradients="FF0000-0000FF;00FF00-FFFF00" --prop transparency=30 --prop categories="$CATS" --prop data="A:60,90,140,180;B:40,70,100,130"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stackedBar --prop title="stacked + serlines=true" --prop serlines=true --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 7: Overlays ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[7] --type shape --prop text="Overlays — referenceline, errbars, gapwidth, overlap, dataTable" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="referenceline=100" --prop legend=none --prop referenceline=100:FF0000:Target --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="errbars=fixedVal:10" --prop legend=none --prop errbars=fixedVal:10 --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="gapwidth=50 + overlap=20" --prop legend=bottom --prop gapwidth=50 --prop overlap=20 --prop categories="$CATS" --prop data="A:60,90,140,180;B:50,75,110,150"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="dataTable=true" --prop legend=bottom --prop dataTable=true --prop categories="$CATS" --prop data="A:60,90,140,180"

# ==================== Slide 8: Presets & per-series control ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[8] --type shape --prop text="Presets & per-series control" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop preset=minimal --prop title="preset=minimal" --prop legend=bottom --prop categories="$CATS" --prop data="A:60,90,140,180;B:50,75,110,150"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar --prop preset=dark --prop title="preset=dark" --prop legend=bottom --prop categories="$CATS" --prop data="A:60,90,140,180;B:50,75,110,150"
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop preset=corporate --prop title="preset=corporate" --prop legend=bottom --prop categories="$CATS" --prop data="A:60,90,140,180;B:50,75,110,150"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar --prop title="seriesN.* Add + chart-series Set" --prop legend=bottom --prop categories="$CATS" --prop series1.name="Product A" --prop series1.values="60,90,140,180" --prop series1.color=4472C4 --prop series2.name="Product B" --prop series2.values="50,75,110,150" --prop series2.color=ED7D31
$CLI set "$FILE" "/slide[8]/chart[4]/series[1]" --prop name=Renamed --prop color=C00000

$CLI close "$FILE"
$CLI validate "$FILE"
echo "Generated: $FILE"
