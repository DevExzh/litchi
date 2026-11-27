#!/bin/bash
# 3D Charts Showcase — column3d / bar3d / pie3d / line3d / area3d with view3d, gapdepth, shape.
# Generates charts-3d.pptx
#
#   Slide 1  3D families            column3d / bar3d / pie3d / line3d
#   Slide 2  area3d & stacked 3D    area3d / stackedColumn3d / percentStackedColumn3d / stackedBar3d
#   Slide 3  view3d                 different rotX,rotY,perspective angles
#   Slide 4  gapdepth               0 / 50 / 150 / 300
#   Slide 5  bar shapes             box / cylinder / cone / pyramid
#   Slide 6  Title & legend
#   Slide 7  Series styling         colors, gradient, transparency, outline, shadow
#   Slide 8  Presets
#
# CLI twin of charts-3d.py (officecli Python SDK). Both produce an equivalent
# charts-3d.pptx.
#
# Usage: ./charts-3d.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/charts-3d.pptx"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ---- shared data + box geometry ----
CATS="Q1,Q2,Q3,Q4"
D2="East:120,135,148,162;West:95,108,115,128"
D3="East:120,135,148,162;South:95,108,115,128;West:80,90,98,110"
PIE_CATS="North,South,East,West"
PIE_D="Share:30,25,28,17"

# title shape helper props are inlined per slide below.

# ==================== Slide 1: 3D families ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[1] --type shape --prop text="3D families — column3d / bar3d / pie3d / line3d" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title=column3d --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop title=bar3d --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[1] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie3d --prop title=pie3d --prop legend=right --prop categories="$PIE_CATS" --prop data="$PIE_D"
$CLI add "$FILE" /slide[1] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=line3d --prop title=line3d --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 2: area3d & stacked 3D ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[2] --type shape --prop text="area3d & stacked 3D" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=area3d --prop title=area3d --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=stackedColumn3d --prop title=stackedColumn3d --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[2] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=percentStackedColumn3d --prop title=percentStackedColumn3d --prop legend=bottom --prop categories="$CATS" --prop data="$D3"
$CLI add "$FILE" /slide[2] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=stackedBar3d --prop title=stackedBar3d --prop legend=bottom --prop categories="$CATS" --prop data="$D3"

# ==================== Slide 3: view3d — rotX,rotY,perspective angles ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[3] --type shape --prop text="view3d — rotX,rotY,perspective angles" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="view3d=15,20,30" --prop view3d="15,20,30" --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="view3d=30,40,15" --prop view3d="30,40,15" --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[3] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="view3d=20" --prop view3d="20" --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[3] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=pie3d --prop title="pie3d view3d=40,30,30" --prop view3d="40,30,30" --prop legend=right --prop categories="$PIE_CATS" --prop data="$PIE_D"

# ==================== Slide 4: gapdepth — 0 / 50 / 150 / 300 ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[4] --type shape --prop text="gapdepth — 0 / 50 / 150 / 300" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="gapdepth=0" --prop gapdepth=0 --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="gapdepth=50" --prop gapdepth=50 --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="gapdepth=150" --prop gapdepth=150 --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[4] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="gapdepth=300" --prop gapdepth=300 --prop legend=none --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 5: 3D bar shapes — box / cylinder / cone / pyramid ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[5] --type shape --prop text="3D bar shapes — box / cylinder / cone / pyramid" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop shape=box --prop title="shape=box" --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop shape=cylinder --prop title="shape=cylinder" --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[5] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop shape=cone --prop title="shape=cone" --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[5] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=bar3d --prop shape=pyramid --prop title="shape=pyramid" --prop legend=none --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 6: Title & legend ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[6] --type shape --prop text="Title & legend" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="Styled title" --prop title.font=Georgia --prop title.size=20 --prop title.color=4472C4 --prop title.bold=true --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="legend=top + legendFont" --prop legend=top --prop legendFont="10:333333:Calibri" --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[6] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="legend.overlay=true" --prop legend=topRight --prop legend.overlay=true --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[6] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop autotitledeleted=true --prop legend=none --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 7: Series styling — colors, gradient, transparency, outline, shadow ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[7] --type shape --prop text="Series styling — colors, gradient, transparency, outline, shadow" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="colors + seriesoutline" --prop colors="4472C4,ED7D31" --prop seriesoutline="000000:0.5" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="gradient + seriesshadow" --prop gradient="FF6600-FFCC00" --prop seriesshadow="000000-5-45-3-50" --prop legend=none --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[7] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="transparency=30" --prop transparency=30 --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[7] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop title="per-series gradients" --prop gradients="FF0000-0000FF;00FF00-FFFF00" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

# ==================== Slide 8: Presets — preset bundles on 3D charts ====================
$CLI add "$FILE" / --type slide
$CLI add "$FILE" /slide[8] --type shape --prop text="Presets — preset bundles on 3D charts" --prop size=24 --prop bold=true --prop autoFit=normal --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop preset=minimal --prop title="preset=minimal" --prop view3d="15,20,30" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=1.05in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop preset=dark --prop title="preset=dark" --prop view3d="15,20,30" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart --prop x=0.3in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop preset=corporate --prop title="preset=corporate" --prop view3d="15,20,30" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"
$CLI add "$FILE" /slide[8] --type chart --prop x=6.95in --prop y=4.25in --prop width=6.1in --prop height=3in --prop chartType=column3d --prop preset=colorful --prop title="preset=colorful" --prop view3d="15,20,30" --prop legend=bottom --prop categories="$CATS" --prop data="$D2"

$CLI close "$FILE"

$CLI validate "$FILE"
echo "Generated: $FILE"
