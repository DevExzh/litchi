#!/bin/bash
# Column Charts Showcase — column, stackedColumn, percentStackedColumn, column3d.
# Generates charts-column.pptx exercising every column-applicable chart property.
# CLI twin of charts-column.py (officecli Python SDK). Both produce an
# equivalent charts-column.pptx.
#
#   Slide 1  Basic variants     column / stackedColumn / percentStackedColumn / column3d
#   Slide 2  Title & legend     title.font/size/color/bold, legend positions, legendFont
#   Slide 3  Data labels        dataLabels flags, labelPos, labelfont
#   Slide 4  Axes               min/max, titles, fonts, gridlines, units, log, secondary
#   Slide 5  Series styling     colors, gradient(s), transparency, outline, shadow, ...
#   Slide 6  Layout & overlays  gapwidth, overlap, referenceline, errbars, trendline, dataTable
#   Slide 7  Backgrounds        chartareafill, plotFill, borders, roundedcorners
#   Slide 8  Presets & per-ser  preset bundles + seriesN.* + chart-series Set
#
# Usage: ./charts-column.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/charts-column.pptx"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# 2x2 grid boxes (widescreen 13.33 x 7.5in): TL / TR / BL / BR
# TL  --prop x=0.3in  --prop y=1.05in --prop width=6.1in --prop height=3in
# TR  --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in
# BL  --prop x=0.3in  --prop y=4.25in --prop width=6.1in --prop height=3in
# BR  --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in
CATS="Q1,Q2,Q3,Q4"
TWO_SERIES="East:120,135,148,162;West:95,108,115,128"
THREE_SERIES="East:120,135,148,162;South:95,108,115,128;West:80,90,98,110"

# ==================== Slide 1 — Basic variants ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[1] --type shape --prop text="Column variants — column / stackedColumn / percentStackedColumn / column3d" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title=column --prop legend=bottom --prop categories="$CATS" --prop data="$TWO_SERIES"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stackedColumn --prop title=stackedColumn --prop legend=bottom --prop categories="$CATS" --prop data="$THREE_SERIES"
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=percentStackedColumn --prop title=percentStackedColumn --prop legend=bottom --prop categories="$CATS" --prop data="$THREE_SERIES"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop view3d=15,20,30 --prop gapdepth=150 --prop title="column3d (view3d=15,20,30)" --prop legend=bottom --prop categories="$CATS" --prop data="$TWO_SERIES"

# ==================== Slide 2 — Title & legend ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[2] --type shape --prop text="Title & legend — title.font/size/color/bold, legend positions, legendFont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop categories="$CATS" --prop data="$TWO_SERIES"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop categories="$CATS" --prop data="$TWO_SERIES"
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="legend=topRight overlay" --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$TWO_SERIES"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop autotitledeleted=true --prop legend=none --prop categories="$CATS" --prop data="$TWO_SERIES"

# ==================== Slide 3 — Data labels ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[3] --type shape --prop text="Data labels — flags (value/category/percent/none), labelPos, labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="value @ outsideEnd" --prop dataLabels=value --prop labelPos=outsideEnd --prop labelfont=10:333333:Calibri --prop legend=none --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="value,category @ insideEnd" --prop dataLabels=value,category --prop labelPos=insideEnd --prop labelfont=9:FFFFFF:Calibri --prop legend=none --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stackedColumn --prop title="stacked + center labels" --prop dataLabels=value --prop labelPos=center --prop labelfont=9:FFFFFF:Calibri --prop legend=bottom --prop categories="$CATS" --prop data="$THREE_SERIES"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="dataLabels=none" --prop dataLabels=none --prop legend=none --prop categories="$CATS" --prop data="A:60,90,140,180"

# ==================== Slide 4 — Axes ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[4] --type shape --prop text="Axes — min/max, titles, fonts, gridlines, units, log, secondary" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="axis min/max + titles + numfmt" --prop legend=none --prop axismin=0 --prop axismax=200 --prop majorunit=50 --prop minorunit=10 --prop axistitle="Revenue (USD)" --prop cattitle=Quarter --prop axisfont=10:333333:Calibri --prop axisline=666666:1 --prop axisnumfmt="#,##0" --prop categories="$CATS" --prop data="Rev:60,90,140,180"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="gridlines + minorGridlines + ticks" --prop legend=none --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop majorTickMark=out --prop minorTickMark=in --prop tickLabelPos=nextTo --prop labelrotation=-30 --prop categories="January,February,March,April" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="dispunits=thousands" --prop legend=none --prop dispunits=thousands --prop categories="$CATS" --prop data="Rev:120000,135000,148000,162000"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=combo --prop combotypes=column,line --prop secondaryaxis=2 --prop title="secondaryaxis=2 (line on right)" --prop legend=bottom --prop categories="$CATS" --prop data="Sales:120,135,148,162;Growth %:5,12,18,22"
# Post-Add chart-axis Set on first chart
$CLI set "$FILE" "/slide[4]/chart[1]/axis[@role=value]" --prop title="Revenue (USD)" --prop format="$#,##0" --prop majorGridlines=true --prop minorGridlines=false --prop max=200 --prop min=0 --prop majorUnit=50

# ==================== Slide 5 — Series styling ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[5] --type shape --prop text="Series styling — colors, gradient(s), transparency, outline, shadow, invertifneg, colorrule" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="colors + seriesoutline" --prop legend=bottom --prop colors=4472C4,ED7D31,A5A5A5 --prop seriesoutline=000000:0.5 --prop categories="$CATS" --prop data="$THREE_SERIES"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="gradient + seriesshadow" --prop legend=bottom --prop gradient=FF6600-FFCC00:90 --prop seriesshadow=000000-5-45-3-50 --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="per-series gradients + transparency=30" --prop legend=bottom --prop gradients="FF0000-0000FF;00FF00-FFFF00" --prop transparency=30 --prop categories="$CATS" --prop data="A:60,90,140,180;B:40,70,100,130"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="invertifneg + colorrule" --prop legend=none --prop invertifneg=true --prop colorrule=0:FF0000:00AA00 --prop categories="Q1,Q2,Q3,Q4,Q5" --prop data="Net:60,-30,40,-50,80"
# Recolor series 1 of the first chart via chart-series Set
$CLI set "$FILE" "/slide[5]/chart[1]/series[1]" --prop color=2E75B6

# ==================== Slide 6 — Layout & overlays ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[6] --type shape --prop text="Layout & overlays — gapwidth, overlap, referenceline, errbars, trendline, dataTable" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="gapwidth=50 + overlap=20" --prop legend=bottom --prop gapwidth=50 --prop overlap=20 --prop categories="$CATS" --prop data="A:60,90,140,180;B:50,75,110,150"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="referenceline=100" --prop legend=none --prop referenceline=100:FF0000:Target --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="errbars=percentage:10" --prop legend=none --prop errbars=percentage:10 --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="dataTable=true + trendline=linear" --prop legend=bottom --prop dataTable=true --prop trendline=linear --prop categories="$CATS" --prop data="A:60,90,140,180"

# ==================== Slide 7 — Backgrounds ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[7] --type shape --prop text="Backgrounds — chartareafill, plotFill, borders, roundedcorners" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="chartareafill + plotFill + borders" --prop legend=bottom --prop chartareafill=FFF8E7 --prop plotFill=FAFAFA --prop chartborder=000000:1 --prop plotborder=CCCCCC:0.5 --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="roundedcorners=true" --prop legend=bottom --prop roundedcorners=true --prop chartborder=4472C4:2 --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="plotFill=none, gridlines=none" --prop legend=none --prop plotFill=none --prop gridlines=none --prop categories="$CATS" --prop data="A:60,90,140,180"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="varyColors=true (single series)" --prop legend=none --prop varyColors=true --prop categories="$CATS" --prop data="A:60,90,140,180"

# ==================== Slide 8 — Presets & per-series control ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[8] --type shape --prop text="Presets & per-series — preset bundles + seriesN.* + chart-series Set" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop preset=minimal --prop title="preset=minimal" --prop legend=bottom --prop categories="$CATS" --prop data="A:60,90,140,180;B:50,75,110,150"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column --prop preset=corporate --prop title="preset=corporate" --prop legend=bottom --prop categories="$CATS" --prop data="A:60,90,140,180;B:50,75,110,150"
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop preset=dark --prop title="preset=dark" --prop legend=bottom --prop categories="$CATS" --prop data="A:60,90,140,180;B:50,75,110,150"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column --prop title="seriesN.* Add + chart-series Set" --prop legend=bottom --prop categories="$CATS" --prop series1.name="Product A" --prop series1.values="60,90,140,180" --prop series1.color=4472C4 --prop series2.name="Product B" --prop series2.values="50,75,110,150" --prop series2.color=ED7D31 --prop series3.name="Product C" --prop series3.values="40,65,90,120" --prop series3.color=70AD47
$CLI set "$FILE" "/slide[8]/chart[4]/series[1]" --prop name="Renamed Alpha" --prop color=C00000

$CLI close "$FILE"
$CLI validate "$FILE"
echo "Generated: $FILE"
