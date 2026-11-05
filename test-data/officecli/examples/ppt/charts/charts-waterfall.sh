#!/bin/bash
# Waterfall Charts Showcase — increaseColor / decreaseColor / totalColor.
# Generates charts-waterfall.pptx
#
#   Slide 1  Basic                  default colors, single dataset
#   Slide 2  Color schemes          increaseColor / decreaseColor / totalColor combinations
#   Slide 3  Title & legend
#   Slide 4  Data labels
#   Slide 5  Axes                   min/max, gridlines, axisnumfmt (currency)
#   Slide 6  Backgrounds            chartareafill, plotFill, chartborder, roundedcorners
#   Slide 7  Larger story           a real cashflow waterfall with labels
#   Slide 8  Presets
#
# CLI twin of charts-waterfall.py (officecli Python SDK).
# Usage: ./charts-waterfall.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-waterfall.pptx"
rm -f "$FILE"

CATS="Start,Q1,Q2,Q3,Q4,End"
D="Cashflow:100,30,-15,40,-10,145"
CATS_LONG="Open,Revenue,COGS,Opex,R&D,Tax,Net"
D_LONG="P&L:100,80,-30,-25,-15,-10,100"

officecli create "$FILE"
officecli open "$FILE"

# ==================== Slide 1: Basic waterfall — default colors ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[1] --type shape --prop text="Basic waterfall — default colors" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="Default colors" --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="Default + dataTable" --prop dataTable=true --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="With legend" --prop legend=bottom --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="7-step P&L" --prop legend=none --prop categories="$CATS_LONG" --prop data="$D_LONG"

# ==================== Slide 2: Color schemes ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[2] --type shape --prop text="Color schemes — increaseColor / decreaseColor / totalColor" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="green/red/blue (default-ish)" --prop increaseColor=00AA00 --prop decreaseColor=FF0000 --prop totalColor=4472C4 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="corporate (teal/orange/navy)" --prop increaseColor=008080 --prop decreaseColor=D86600 --prop totalColor=1F3864 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="monochrome" --prop increaseColor=606060 --prop decreaseColor=A0A0A0 --prop totalColor=303030 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="vivid" --prop increaseColor=00C853 --prop decreaseColor=D50000 --prop totalColor=2962FF --prop legend=none --prop categories="$CATS" --prop data="$D"

# ==================== Slide 3: Title & legend ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[3] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop autotitledeleted=true --prop legend=none --prop categories="$CATS" --prop data="$D"

# ==================== Slide 4: Data labels ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[4] --type shape --prop text="Data labels — flags + labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="value" --prop dataLabels=value --prop labelfont=10:333333:Calibri --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="value,category" --prop dataLabels=value,category --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="value @ outsideEnd" --prop dataLabels=value --prop labelPos=outsideEnd --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="dataLabels=none" --prop dataLabels=none --prop legend=none --prop categories="$CATS" --prop data="$D"

# ==================== Slide 5: Axes ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[5] --type shape --prop text="Axes — min/max, titles, gridlines, axisnumfmt" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="min/max + titles" --prop axismin=0 --prop axismax=200 --prop majorunit=50 --prop axistitle=USD --prop cattitle=Phase --prop axisfont=10:333333:Calibri --prop axisnumfmt='$#,##0' --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="gridlines + minorGridlines" --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="labelrotation=-30" --prop labelrotation=-30 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="dispunits=thousands" --prop dispunits=thousands --prop legend=none --prop categories="$CATS" --prop data="USD:100000,30000,-15000,40000,-10000,145000"

# ==================== Slide 6: Backgrounds ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[6] --type shape --prop text="Backgrounds — chartareafill, plotFill, chartborder, roundedcorners" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="chartareafill + chartborder" --prop chartareafill=FFF8E7 --prop chartborder=000000:1 --prop plotFill=FAFAFA --prop plotborder=CCCCCC:0.5 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="roundedcorners=true" --prop roundedcorners=true --prop chartborder=4472C4:2 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="plotFill=none" --prop plotFill=none --prop gridlines=none --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop title="chartareafill=none" --prop chartareafill=none --prop legend=none --prop categories="$CATS" --prop data="$D"

# ==================== Slide 7: Hero cashflow waterfall ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[7] --type shape --prop text="Hero cashflow waterfall — full slide with labels" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[7] --type chart --prop x=1in --prop y=1.05in --prop width=11.3in --prop height=6.2in --prop chartType=waterfall --prop title="FY24 P&L Walk" --prop title.font=Helvetica --prop title.size=22 --prop title.bold=true --prop title.color=1F3864 --prop increaseColor=00C853 --prop decreaseColor=D50000 --prop totalColor=2962FF --prop dataLabels=value,category --prop labelPos=outsideEnd --prop labelfont=11:333333:Helvetica --prop axistitle=USD --prop cattitle= --prop axisnumfmt='$#,##0' --prop gridlines=E0E0E0:0.3 --prop legend=none --prop categories="$CATS_LONG" --prop data="$D_LONG"

# ==================== Slide 8: Presets ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[8] --type shape --prop text="Presets" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop preset=minimal --prop title="preset=minimal" --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop preset=dark --prop title="preset=dark" --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop preset=corporate --prop title="preset=corporate" --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=waterfall --prop preset=colorful --prop title="preset=colorful" --prop legend=none --prop categories="$CATS" --prop data="$D"

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
