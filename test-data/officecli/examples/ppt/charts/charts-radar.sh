#!/bin/bash
# Radar Charts Showcase — radarstyle standard / marker / filled.
# CLI twin of charts-radar.py (officecli Python SDK). Both produce an
# equivalent charts-radar.pptx.
#
#   Slide 1  radarstyle             standard / marker / filled
#   Slide 2  Title & legend         title.* + legend positions + legendFont
#   Slide 3  Data labels            flags + labelfont
#   Slide 4  Axes                   min/max, gridlines, axisfont, labelrotation
#   Slide 5  Series styling         colors, gradient, transparency, outline, shadow
#   Slide 6  Markers                marker symbol/size/color (radarstyle=marker only)
#   Slide 7  Backgrounds            chartareafill, plotFill, chartborder, roundedcorners
#   Slide 8  Presets & per-series   preset bundles + chart-series Set
#
# Usage: ./charts-radar.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-radar.pptx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# Shared category/data series and four-up grid boxes (inches).
CATS="Speed,Power,Range,Style,Tech,Price"
D="A:8,7,9,6,8,7"
D2="Model A:8,7,9,6,8,7;Model B:6,9,7,8,9,6"

# ==================== Slide 1: radarstyle — standard / marker / filled ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[1] --type shape --prop text="radarstyle — standard / marker / filled" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title=radarstyle=standard --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=marker --prop title=radarstyle=marker --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title=radarstyle=filled --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title="single series" --prop legend=bottom --prop categories="$CATS" --prop data="$D"

# ==================== Slide 2: Title & legend ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[2] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title="legend=top + legendFont" --prop legend=top --prop legendFont=10:333333:Calibri --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop autotitledeleted=true --prop legend=none --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 3: Data labels — flags + labelfont ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[3] --type shape --prop text="Data labels — flags + labelfont" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=marker --prop title=value --prop dataLabels=value --prop labelfont=9:333333:Calibri --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=marker --prop title="value,series" --prop dataLabels="value,series" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title="value,category" --prop dataLabels="value,category" --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title="dataLabels=none" --prop dataLabels=none --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 4: Axes — min/max, gridlines, axisfont, labelrotation ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[4] --type shape --prop text="Axes — min/max, gridlines, axisfont, labelrotation" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title="min/max + titles" --prop axismin=0 --prop axismax=10 --prop majorunit=2 --prop axisfont=10:333333:Calibri --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title="gridlines + minorGridlines" --prop gridlines=E0E0E0:0.3 --prop minorGridlines=F0F0F0:0.25 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title="labelrotation=30" --prop labelrotation=30 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title="axisnumfmt=0.0" --prop axisnumfmt=0.0 --prop legend=none --prop categories="$CATS" --prop data="$D"

# ==================== Slide 5: Series styling — colors, gradient, transparency, outline, shadow ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[5] --type shape --prop text="Series styling — colors, gradient, transparency, outline, shadow" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title="colors + seriesoutline" --prop colors=4472C4,ED7D31 --prop seriesoutline=000000:0.5 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title="gradient + seriesshadow" --prop gradient=FF6600-FFCC00 --prop seriesshadow=000000-5-45-3-50 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title="transparency=40" --prop transparency=40 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title="per-series gradients" --prop gradients="FF0000-0000FF;00FF00-FFFF00" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 6: Markers (radarstyle=marker) — symbol/size/color ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[6] --type shape --prop text="Markers (radarstyle=marker) — symbol/size/color" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=marker --prop title=circle:10:FF0000 --prop marker=circle:10:FF0000 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=marker --prop title=square:8:0070C0 --prop marker=square:8:0070C0 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=marker --prop title=diamond:12 --prop marker=diamond:12 --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=marker --prop title=triangle:10:70AD47 --prop marker=triangle:10:70AD47 --prop legend=none --prop categories="$CATS" --prop data="$D"

# ==================== Slide 7: Backgrounds — chartareafill, plotFill, chartborder, roundedcorners ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[7] --type shape --prop text="Backgrounds — chartareafill, plotFill, chartborder, roundedcorners" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title="chartareafill + plotFill + borders" --prop chartareafill=FFF8E7 --prop plotFill=FAFAFA --prop chartborder=000000:1 --prop plotborder=CCCCCC:0.5 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title="roundedcorners=true" --prop roundedcorners=true --prop chartborder=4472C4:2 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=standard --prop title="plotFill=none" --prop plotFill=none --prop legend=none --prop categories="$CATS" --prop data="$D"
officecli add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop title="chartareafill=none" --prop chartareafill=none --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 8: Presets & per-series Set ====================
officecli add "$FILE" / --type slide
officecli add "$FILE" /slide[8] --type shape --prop text="Presets & per-series Set" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
officecli add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop preset=minimal --prop title=preset=minimal --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop preset=dark --prop title=preset=dark --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=filled --prop preset=corporate --prop title=preset=corporate --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
officecli add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=radar --prop radarstyle=marker --prop title="chart-series Set" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
# per-series Set: recolor + remark the first series of chart[4]
officecli set "$FILE" /slide[8]/chart[4]/series[1] --prop name="Renamed A" --prop color=C00000 --prop marker=circle --prop markerSize=9

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
