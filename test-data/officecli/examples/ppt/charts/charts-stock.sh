#!/bin/bash
# Stock Charts Showcase — High-Low-Close and OHLC variants.
# Generates: charts-stock.pptx
#
#   Slide 1  Basic stock         3-series HLC + 4-series OHLC
#   Slide 2  Hi-low / up-down    hilowlines, updownbars
#   Slide 3  Title & legend
#   Slide 4  Data labels
#   Slide 5  Axes                min/max, gridlines, axisnumfmt (currency)
#   Slide 6  Series styling      colors, transparency, outline, shadow
#   Slide 7  Backgrounds         chartareafill, plotFill, chartborder
#   Slide 8  Presets & per-ser   preset bundles + chart-series Set
#
# CLI twin of charts-stock.py (officecli Python SDK).
# Usage: ./charts-stock.sh [officecli path]

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/charts-stock.pptx"

# shared geometry + data
CATS="Mon,Tue,Wed,Thu,Fri"
HLC="High:130,135,140,138,145;Low:118,122,128,125,132;Close:125,130,135,132,140"
OHLC="Open:120,128,130,135,138;High:130,135,140,138,145;Low:118,122,128,125,132;Close:125,130,135,132,140"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ==================== Slide 1: Basic stock — HLC vs OHLC ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[1] --type shape --prop text="Basic stock — High-Low-Close vs Open-High-Low-Close" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title=HLC --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title=OHLC --prop legend=bottom --prop categories="$CATS" --prop data="$OHLC"
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="HLC + dataTable=true" --prop dataTable=true --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="OHLC + dataTable=true" --prop dataTable=true --prop legend=bottom --prop categories="$CATS" --prop data="$OHLC"

# ==================== Slide 2: hilowlines & updownbars ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[2] --type shape --prop text="hilowlines & updownbars" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="hilowlines=true" --prop hilowlines=true --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="hilowlines=808080:0.5" --prop hilowlines=808080:0.5 --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="updownbars=true (OHLC)" --prop updownbars=true --prop legend=bottom --prop categories="$CATS" --prop data="$OHLC"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="updownbars=150:00AA00:FF0000" --prop updownbars=150:00AA00:FF0000 --prop legend=bottom --prop categories="$CATS" --prop data="$OHLC"

# ==================== Slide 3: Title & legend ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[3] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop autotitledeleted=true --prop legend=none --prop categories="$CATS" --prop data="$HLC"

# ==================== Slide 4: Data labels — flags + labelfont ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[4] --type shape --prop text="Data labels — flags + labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="dataLabels=value" --prop dataLabels=value --prop labelfont=9:333333:Calibri --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="value,series" --prop dataLabels=value,series --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="value,category" --prop dataLabels=value,category --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="dataLabels=none" --prop dataLabels=none --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"

# ==================== Slide 5: Axes — min/max, gridlines, currency format ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[5] --type shape --prop text="Axes — min/max, gridlines, currency format" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="min/max + titles" --prop axismin=100 --prop axismax=160 --prop majorunit=10 --prop axistitle="Price (USD)" --prop cattitle=Day --prop axisfont=10:333333:Calibri --prop axisnumfmt="\$#,##0.00" --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="gridlines + minorGridlines" --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="labelrotation=-30" --prop labelrotation=-30 --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="dispunits=hundreds" --prop dispunits=hundreds --prop legend=bottom --prop categories="$CATS" --prop data="High:13000,13500,14000,13800,14500;Low:11800,12200,12800,12500,13200;Close:12500,13000,13500,13200,14000"

# ==================== Slide 6: Series styling — colors, transparency, outline, shadow ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[6] --type shape --prop text="Series styling — colors, transparency, outline, shadow" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="colors" --prop colors=4472C4,ED7D31,70AD47 --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="seriesoutline" --prop seriesoutline=000000:1 --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="transparency=30" --prop transparency=30 --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="seriesshadow" --prop seriesshadow=000000-5-45-3-50 --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"

# ==================== Slide 7: Backgrounds — chartareafill, plotFill, chartborder, roundedcorners ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[7] --type shape --prop text="Backgrounds — chartareafill, plotFill, chartborder, roundedcorners" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="chartareafill + plotFill + borders" --prop chartareafill=FFF8E7 --prop plotFill=FAFAFA --prop chartborder=000000:1 --prop plotborder=CCCCCC:0.5 --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="roundedcorners=true" --prop roundedcorners=true --prop chartborder=4472C4:2 --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="plotFill=none" --prop plotFill=none --prop gridlines=none --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="chartareafill=none" --prop chartareafill=none --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"

# ==================== Slide 8: Presets & per-series Set ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[8] --type shape --prop text="Presets & per-series Set" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop preset=minimal --prop title="preset=minimal" --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stock --prop preset=dark --prop title="preset=dark" --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop preset=corporate --prop title="preset=corporate" --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stock --prop title="chart-series Set name+color" --prop legend=bottom --prop categories="$CATS" --prop data="$HLC"
$CLI set "$FILE" "/slide[8]/chart[4]/series[1]" --prop name=H --prop color=00AA00
$CLI set "$FILE" "/slide[8]/chart[4]/series[2]" --prop name=L --prop color=C00000
$CLI set "$FILE" "/slide[8]/chart[4]/series[3]" --prop name=C --prop color=4472C4

$CLI close "$FILE"
$CLI validate "$FILE"
echo "Generated: $FILE"
