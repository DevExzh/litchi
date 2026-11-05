#!/bin/bash
# Pie & Doughnut Charts Showcase — pie, pie3d, and doughnut with all variations.
# Generates: charts-pie.xlsx  (3 chart sheets, 12 charts total)
#
# CLI twin of charts-pie.py (officecli Python SDK). Both produce an
# equivalent charts-pie.xlsx.
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-pie.xlsx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Sheet: 1-Pie Charts
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="1-Pie Charts"

# Chart 1: Basic pie chart with inline data and custom colors
# Features: chartType=pie, inline series, categories, colors, dataLabels
officecli add "$FILE" "/1-Pie Charts" --type chart \
  --prop chartType=pie \
  --prop title="Market Share" \
  --prop series1=Share:40,25,20,15 \
  --prop categories=Product A,Product B,Product C,Product D \
  --prop colors=4472C4,ED7D31,70AD47,FFC000 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=outsideEnd

# Chart 2: Pie with exploded slice and per-point colors
# Features: explosion (slice separation %), point{N}.color,
#   labelPos=bestFit, dataLabels=percent
officecli add "$FILE" "/1-Pie Charts" --type chart \
  --prop chartType=pie \
  --prop title="Revenue by Region" \
  --prop series1=Revenue:35,28,22,15 \
  --prop categories=North,South,East,West \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop explosion=15 \
  --prop point1.color=1F4E79 --prop point2.color=2E75B6 \
  --prop point3.color=9DC3E6 --prop point4.color=BDD7EE \
  --prop dataLabels=percent --prop labelPos=bestFit

# Chart 3: 3D pie with perspective and title styling
# Features: pie3d, view3d on pie (tilt angle), title.font/size/color/bold,
#   labelFont (size:color:bold)
officecli add "$FILE" "/1-Pie Charts" --type chart \
  --prop chartType=pie3d \
  --prop title="3D Category Split" \
  --prop series1=Sales:45,30,25 \
  --prop categories=Electronics,Clothing,Food \
  --prop colors=2E75B6,70AD47,FFC000 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop view3d=30,0,0 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop dataLabels=true --prop labelPos=center \
  --prop labelFont=12:FFFFFF:true

# Chart 4: Pie with gradient fills, leader lines, and legend positioning
# Features: gradients (per-slice), legend=right, legendfont,
#   dataLabels.showLeaderLines, chartFill, roundedCorners
officecli add "$FILE" "/1-Pie Charts" --type chart \
  --prop chartType=pie \
  --prop title="Q4 Distribution" \
  --prop series1=Q4:198,158,142,180 \
  --prop categories=East,South,North,West \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop 'gradients=4472C4-BDD7EE:90;ED7D31-FBE5D6:90;70AD47-C5E0B4:90;FFC000-FFF2CC:90' \
  --prop legend=right --prop legendfont=10:333333:Helvetica \
  --prop dataLabels=true \
  --prop dataLabels.showLeaderLines=true \
  --prop chartFill=FAFAFA --prop roundedCorners=true

# ==========================================================================
# Sheet: 2-Doughnut Charts
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="2-Doughnut Charts"

# Chart 1: Basic doughnut chart
# Features: chartType=doughnut, center labels
officecli add "$FILE" "/2-Doughnut Charts" --type chart \
  --prop chartType=doughnut \
  --prop title="Channel Mix" \
  --prop series1=Channel:55,45 \
  --prop categories=Online,Retail \
  --prop colors=4472C4,ED7D31 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=center \
  --prop labelFont=14:FFFFFF:true

# Chart 2: Multi-ring doughnut (multiple series)
# Features: multi-ring doughnut (multiple series = concentric rings),
#   series.outline (white separator between slices)
officecli add "$FILE" "/2-Doughnut Charts" --type chart \
  --prop chartType=doughnut \
  --prop title="Year-over-Year Comparison" \
  --prop series1=2024:40,35,25 \
  --prop series2=2025:45,30,25 \
  --prop categories=Electronics,Clothing,Food \
  --prop colors=4472C4,70AD47,FFC000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop series.outline=FFFFFF-1 \
  --prop legend=bottom

# Chart 3: Styled doughnut with shadow and custom data labels
# Features: series.shadow on doughnut, title.shadow, plotFill
officecli add "$FILE" "/2-Doughnut Charts" --type chart \
  --prop chartType=doughnut \
  --prop title="Priority Breakdown" \
  --prop series1=Priority:50,30,20 \
  --prop categories=High,Medium,Low \
  --prop colors=C00000,FFC000,70AD47 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop series.shadow=000000-4-315-2-30 \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop 'dataLabels.numFmt=0"%"' \
  --prop title.shadow=000000-3-315-2-30 \
  --prop plotFill=F5F5F5

# Chart 4: Doughnut with per-slice gradient and explosion
# Features: explosion on doughnut, 5-slice gradients
officecli add "$FILE" "/2-Doughnut Charts" --type chart \
  --prop chartType=doughnut \
  --prop title="Product Revenue" \
  --prop series1=Revenue:35,25,20,12,8 \
  --prop categories=Laptop,Phone,Tablet,Jacket,Coffee \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop explosion=8 \
  --prop 'gradients=1F4E79-5B9BD5:90;C55A11-F4B183:90;548235-A9D18E:90;7F6000-FFD966:90;843C0B-DDA15E:90' \
  --prop legend=right \
  --prop dataLabels=true --prop labelPos=bestFit

# ==========================================================================
# Sheet: 3-Pie Advanced
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="3-Pie Advanced"

# Chart 1: varyColors=true + firstSliceAngle on pie
# Features: varyColors=true (each slice gets a distinct color automatically),
#   firstSliceAngle=45 (rotates the first slice start angle, 0-360 degrees)
officecli add "$FILE" "/3-Pie Advanced" --type chart \
  --prop chartType=pie \
  --prop title="Pie — varyColors + firstSliceAngle" \
  --prop series1=Share:40,30,20,10 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop varyColors=true \
  --prop firstSliceAngle=45 \
  --prop dataLabels=true --prop labelPos=bestFit

# Chart 2: holeSize + leaderlines on doughnut
# Features: holeSize=65 (% of total radius — larger value = thinner ring),
#   leaderlines=true (connecting lines from labels to slices, pie/doughnut)
officecli add "$FILE" "/3-Pie Advanced" --type chart \
  --prop chartType=doughnut \
  --prop title="Doughnut — holeSize + leaderlines" \
  --prop series1=Revenue:35,28,22,15 \
  --prop categories=North,South,East,West \
  --prop colors=2E75B6,ED7D31,70AD47,FFC000 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop holeSize=65 \
  --prop leaderlines=true \
  --prop dataLabels=true --prop labelPos=outsideEnd

# Chart 3: title.overlay on pie (title floats over plot area)
# Features: title.overlay=true (title overlays the plot area, maximizing
#   chart area — contrast with default where title reserves space above)
officecli add "$FILE" "/3-Pie Advanced" --type chart \
  --prop chartType=pie \
  --prop title="Overlaid Title" \
  --prop title.overlay=true \
  --prop series1=Mix:50,30,20 \
  --prop categories=Online,Retail,Partner \
  --prop colors=4472C4,70AD47,FFC000 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop varyColors=false \
  --prop dataLabels=percent --prop labelPos=center

# Chart 4: Doughnut — holeSize + firstSliceAngle + title.overlay combined
# Features: three doughnut-specific props together —
#   holeSize, varyColors, title.overlay
officecli add "$FILE" "/3-Pie Advanced" --type chart \
  --prop chartType=doughnut \
  --prop title="Doughnut — Combined" \
  --prop title.overlay=true \
  --prop series1=Split:45,35,20 \
  --prop categories=A,B,C \
  --prop colors=C00000,FFC000,548235 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop holeSize=50 \
  --prop varyColors=false \
  --prop dataLabels=true --prop labelPos=center \
  --prop labelFont=12:FFFFFF:true

# Remove blank default Sheet1 (all data is inline)
officecli remove "$FILE" /Sheet1

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
