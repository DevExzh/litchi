#!/bin/bash
# Radar Charts Showcase — radar with standard, filled, and marker styles.
# Generates charts-radar.xlsx (16 charts across 4 sheets).
# CLI twin of charts-radar.py (officecli Python SDK).
# Usage: ./charts-radar.sh

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-radar.xlsx"

rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Sheet: 1-Radar Fundamentals
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="1-Radar Fundamentals"

# Chart 1: Basic radar (standard style) with 3 series
officecli add "$FILE" "/1-Radar Fundamentals" --type chart \
  --prop chartType=radar \
  --prop radarStyle=standard \
  --prop title="Athlete Comparison" \
  --prop series1="Alice:85,70,90,60,75" \
  --prop series2="Bob:65,90,70,80,85" \
  --prop series3="Carol:75,80,80,70,65" \
  --prop categories=Speed,Strength,Stamina,Agility,Accuracy \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop legend=bottom

# Chart 2: Radar with markers (marker style)
officecli add "$FILE" "/1-Radar Fundamentals" --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop title="Product Ratings" \
  --prop series1="Product A:9,7,8,6,8" \
  --prop series2="Product B:6,9,7,8,5" \
  --prop categories=Quality,Price,Design,Support,Delivery \
  --prop colors=2E75B6,C00000 \
  --prop marker=circle:6:2E75B6 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop legend=bottom \
  --prop dataLabels=true

# Chart 3: Filled radar with transparency
officecli add "$FILE" "/1-Radar Fundamentals" --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop title="Skills Assessment" \
  --prop series1="Junior:50,40,60,70,55" \
  --prop series2="Senior:85,80,75,90,80" \
  --prop categories=Coding,Design,Testing,Communication,Leadership \
  --prop colors=4472C4,70AD47 \
  --prop transparency=40 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom

# Chart 4: Filled radar with per-series colors and white outline
officecli add "$FILE" "/1-Radar Fundamentals" --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop title="Department Scores" \
  --prop series1="Engineering:90,75,60,85,70" \
  --prop series2="Marketing:60,85,80,70,90" \
  --prop series3="Sales:70,80,75,65,85" \
  --prop categories=Innovation,Teamwork,Efficiency,Quality,Growth \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop series.outline=FFFFFF-0.5 \
  --prop transparency=35 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom

# ==========================================================================
# Sheet: 2-Radar Styling
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="2-Radar Styling"

# Chart 1: Title styling (font, size, color, bold, shadow)
officecli add "$FILE" "/2-Radar Styling" --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop title="Styled Title Demo" \
  --prop series1="Team A:80,65,90,70,85" \
  --prop categories=Attack,Defense,Speed,Skill,Stamina \
  --prop colors=2E75B6 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop title.font=Georgia --prop title.size=18 \
  --prop title.color=1F4E79 --prop title.bold=true \
  --prop title.shadow=000000-3-315-2-30

# Chart 2: Series shadow effects
officecli add "$FILE" "/2-Radar Styling" --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop title="Shadow Effects" \
  --prop series1="Region A:75,80,65,90,70" \
  --prop series2="Region B:60,70,85,75,80" \
  --prop categories=Revenue,Profit,Growth,Retention,Satisfaction \
  --prop colors=4472C4,ED7D31 \
  --prop series.shadow=000000-4-315-2-30 \
  --prop transparency=30 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop legend=bottom

# Chart 3: Axis font and gridlines styling
officecli add "$FILE" "/2-Radar Styling" --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop title="Axis & Gridlines" \
  --prop series1="Actual:70,85,60,75,80" \
  --prop series2="Target:80,80,80,80,80" \
  --prop categories="KPI 1,KPI 2,KPI 3,KPI 4,KPI 5" \
  --prop colors=4472C4,C00000 \
  --prop axisfont=10:333333:Calibri \
  --prop gridlines=D9D9D9:0.5 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom

# Chart 4: Plot fill, chart fill, rounded corners, borders
officecli add "$FILE" "/2-Radar Styling" --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop title="Chart Area Styling" \
  --prop series1="Score:85,70,90,60,75" \
  --prop categories=Speed,Power,Technique,Endurance,Flexibility \
  --prop colors=4472C4 \
  --prop transparency=25 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop plotFill=F5F5F5 --prop chartFill=FAFAFA \
  --prop roundedCorners=true \
  --prop chartArea.border=BFBFBF:0.5 \
  --prop plotArea.border=D9D9D9:0.25

# ==========================================================================
# Sheet: 3-Labels & Legend
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="3-Labels & Legend"

# Chart 1: Data labels with font styling and position
officecli add "$FILE" "/3-Labels & Legend" --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop title="Data Labels" \
  --prop series1="Performance:88,72,95,67,81" \
  --prop categories=Speed,Strength,Stamina,Agility,Accuracy \
  --prop colors=2E75B6 \
  --prop marker=circle:6:2E75B6 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelPos=outsideEnd \
  --prop labelFont=9:333333:true

# Chart 2: Legend positioning and styling with overlay
officecli add "$FILE" "/3-Labels & Legend" --type chart \
  --prop chartType=radar \
  --prop radarStyle=standard \
  --prop title="Legend Styles" \
  --prop series1="Alpha:80,60,75,90,70" \
  --prop series2="Beta:70,80,85,65,75" \
  --prop series3="Gamma:65,75,70,80,85" \
  --prop categories="Metric A,Metric B,Metric C,Metric D,Metric E" \
  --prop colors=4472C4,ED7D31,70AD47 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop legend=right \
  --prop legendfont=10:1F4E79:Calibri \
  --prop legend.overlay=true

# Chart 3: Manual plot area layout
officecli add "$FILE" "/3-Labels & Legend" --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop title="Custom Layout" \
  --prop series1="Team:85,70,90,65,80" \
  --prop categories=Vision,Execution,Culture,Agility,Impact \
  --prop colors=4472C4 \
  --prop transparency=30 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop plotArea.x=0.1 --prop plotArea.y=0.15 \
  --prop plotArea.w=0.8 --prop plotArea.h=0.75

# Chart 4: Multiple series (5+) comparison
officecli add "$FILE" "/3-Labels & Legend" --type chart \
  --prop chartType=radar \
  --prop radarStyle=standard \
  --prop title="Multi-Team Comparison" \
  --prop series1="Dev:90,70,80,65,75" \
  --prop series2="QA:60,85,70,80,90" \
  --prop series3="Design:75,80,85,70,60" \
  --prop series4="PM:80,65,75,90,70" \
  --prop series5="DevOps:70,75,60,85,80" \
  --prop categories=Speed,Quality,Innovation,Teamwork,Delivery \
  --prop colors=4472C4,ED7D31,70AD47,FFC000,7030A0 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom \
  --prop legendfont=9:333333:Calibri

# ==========================================================================
# Sheet: 4-Advanced
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="4-Advanced"

# Chart 1: Title glow and shadow effects
officecli add "$FILE" "/4-Advanced" --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop title="Glow & Shadow Title" \
  --prop series1="Score:75,85,65,90,80" \
  --prop categories=Creativity,Logic,Memory,Focus,Speed \
  --prop colors=2E75B6 \
  --prop marker=diamond:7:2E75B6 \
  --prop x=0 --prop y=0 --prop width=12 --prop height=18 \
  --prop title.font=Georgia --prop title.size=16 \
  --prop title.bold=true --prop title.color=1F4E79 \
  --prop title.glow=4472C4-8 \
  --prop title.shadow=000000-3-315-2-30

# Chart 2: Radar with many spokes (8 categories)
officecli add "$FILE" "/4-Advanced" --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop title="8-Spoke Assessment" \
  --prop series1="Candidate:85,70,90,60,75,80,65,88" \
  --prop series2="Benchmark:70,70,70,70,70,70,70,70" \
  --prop categories=Technical,Communication,Leadership,Creativity,Analytical,Teamwork,Adaptability,Initiative \
  --prop colors=4472C4,BFBFBF \
  --prop transparency=35 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop legend=bottom \
  --prop gridlines=D9D9D9:0.25

# Chart 3: Single-series radar with full styling
officecli add "$FILE" "/4-Advanced" --type chart \
  --prop chartType=radar \
  --prop radarStyle=marker \
  --prop title="Personal Profile" \
  --prop series1="Self:92,78,85,65,88,70" \
  --prop categories=Python,JavaScript,SQL,DevOps,Testing,Design \
  --prop colors=7030A0 \
  --prop marker=square:7:7030A0 \
  --prop x=0 --prop y=19 --prop width=12 --prop height=18 \
  --prop dataLabels=true --prop labelFont=9:7030A0:true \
  --prop title.font=Calibri --prop title.size=14 \
  --prop title.color=7030A0 --prop title.bold=true \
  --prop plotFill=F8F0FF --prop chartFill=FFFFFF \
  --prop roundedCorners=true \
  --prop chartArea.border=7030A0:0.5

# Chart 4: Two-series filled radar with low transparency for overlap
officecli add "$FILE" "/4-Advanced" --type chart \
  --prop chartType=radar \
  --prop radarStyle=filled \
  --prop title="Before vs After" \
  --prop series1="Before:55,40,65,50,45" \
  --prop series2="After:85,75,80,70,80" \
  --prop categories=Revenue,Efficiency,Satisfaction,Innovation,Retention \
  --prop colors=C00000,70AD47 \
  --prop transparency=20 \
  --prop series.outline=FFFFFF-0.75 \
  --prop x=13 --prop y=19 --prop width=12 --prop height=18 \
  --prop legend=bottom \
  --prop dataLabels=true --prop labelFont=9:333333:false \
  --prop chartFill=FAFAFA --prop plotFill=F5F5F5

# Remove blank default Sheet1 (all data is inline)
officecli remove "$FILE" /Sheet1

officecli close "$FILE"

officecli validate "$FILE"
echo "Generated: $FILE"
