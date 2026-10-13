#!/bin/bash
# Extended Chart Types Showcase — full feature coverage for waterfall, funnel,
# treemap, sunburst, histogram, boxWhisker (cx:chart family) plus pareto and
# chart-meta knobs (anchor, preset, autotitledeleted, plotvisonly).
#
# CLI twin of charts-extended.py (officecli Python SDK). Both produce an
# equivalent charts-extended.xlsx.
#
# Usage:
#   ./charts-extended.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-extended.xlsx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Sheet 1: Waterfall & Funnel
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="1-Waterfall & Funnel"

# Chart 1: Waterfall — increase/decrease/total colors + data labels + title glow
# Features: chartType=waterfall, increaseColor, decreaseColor, totalColor,
#   dataLabels, title.glow
officecli add "$FILE" "/1-Waterfall & Funnel" --type chart \
  --prop chartType=waterfall \
  --prop title="Cash Flow Bridge" \
  --prop data=Start:1000,Revenue:500,Costs:-300,Tax:-100,Net:1100 \
  --prop increaseColor=70AD47 \
  --prop decreaseColor=FF0000 \
  --prop totalColor=4472C4 \
  --prop dataLabels=true \
  --prop title.glow=00D2FF-6-60 \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# Chart 2: Waterfall — chart-area fill + legend + custom label font
# Features: waterfall with legend=bottom, chartFill (solid hex — cx charts
#   don't support gradient fills, use plain RGB), labelFont "size:color:bold"
officecli add "$FILE" "/1-Waterfall & Funnel" --type chart \
  --prop chartType=waterfall \
  --prop title="Budget vs Actual" \
  --prop data=Budget:5000,Sales:2000,Marketing:-800,Ops:-600,Net:5600 \
  --prop increaseColor=2E75B6 \
  --prop decreaseColor=C00000 \
  --prop totalColor=FFC000 \
  --prop legend=bottom \
  --prop chartFill=F0F4FA \
  --prop dataLabels=true \
  --prop labelFont=9:333333:true \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# Chart 3: Funnel — sales pipeline with title shadow
# Features: chartType=funnel, descending pipeline values, dataLabels,
#   title.shadow "COLOR-BLUR-ANGLE-DIST-OPACITY"
officecli add "$FILE" "/1-Waterfall & Funnel" --type chart \
  --prop chartType=funnel \
  --prop title="Sales Pipeline" \
  --prop series1=Pipeline:1200,850,600,300,120 \
  --prop categories=Leads,Qualified,Proposal,Negotiation,Won \
  --prop dataLabels=true \
  --prop title.shadow=000000-4-45-2-40 \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# Chart 4: Funnel — marketing conversion + legend/axis fonts
# Features: funnel, legendfont "size:color:fontname", axisfont, 6-stage, dataLabels
# NOTE: `colors=` palette is intentionally omitted. On cx:chart single-series
#   types (funnel/treemap/sunburst) the CLI only applies the first palette color
#   to the whole series, so all bars would render the same color. Let Excel's
#   theme pick the default accent color.
officecli add "$FILE" "/1-Waterfall & Funnel" --type chart \
  --prop chartType=funnel \
  --prop title="Marketing Funnel" \
  --prop series1=Users:10000,6500,3200,1800,900,450 \
  --prop categories=Impressions,Clicks,Signups,Active,Paying,Retained \
  --prop dataLabels=true \
  --prop "legendfont=9:8B949E:Helvetica Neue" \
  --prop "axisfont=10:58626E:Helvetica Neue" \
  --prop x=14 --prop y=19 --prop width=13 --prop height=18

# ==========================================================================
# Sheet 2: Treemap & Sunburst
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="2-Treemap & Sunburst"

# Chart 1: Treemap — parentLabelLayout=overlapping + dataLabels
# Features: chartType=treemap, parentLabelLayout=overlapping, dataLabels.
#   NOTE: `colors=` is omitted — see Funnel Chart 4 note: cx single-series
#   charts only pick up the first palette color. Excel's theme auto-rainbows.
officecli add "$FILE" "/2-Treemap & Sunburst" --type chart \
  --prop chartType=treemap \
  --prop title="Revenue by Product" \
  --prop series1=Revenue:450,380,310,280,210,180,150,120 \
  --prop categories=Laptops,Phones,Tablets,TVs,Cameras,Audio,Gaming,Wearables \
  --prop parentLabelLayout=overlapping \
  --prop dataLabels=true \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# Chart 2: Treemap — parentLabelLayout=banner + bold title
# Features: treemap parentLabelLayout=banner, title.bold/size/color
officecli add "$FILE" "/2-Treemap & Sunburst" --type chart \
  --prop chartType=treemap \
  --prop title="Department Budget" \
  --prop series1=Budget:900,750,600,500,420,350,280 \
  --prop categories=Engineering,Sales,Marketing,Support,Finance,HR,Legal \
  --prop parentLabelLayout=banner \
  --prop title.bold=true \
  --prop title.size=14 \
  --prop title.color=2E5090 \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# Chart 3: Treemap — parentLabelLayout=none (no parent label strip)
# Features: treemap parentLabelLayout=none (all labels inline, no header strip),
#   dataLabels on leaf tiles
officecli add "$FILE" "/2-Treemap & Sunburst" --type chart \
  --prop chartType=treemap \
  --prop title="Flat Treemap (no parent labels)" \
  --prop series1=Units:250,200,180,160,140,120,100,80,60,40 \
  --prop categories=A,B,C,D,E,F,G,H,I,J \
  --prop parentLabelLayout=none \
  --prop dataLabels=true \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# Chart 4: Sunburst — radial hierarchy + chartFill (solid) + plotFill
# Features: chartType=sunburst, radial hierarchical layout, chartFill (solid hex),
#   plotFill (solid hex), dataLabels.
#   NOTE 1: cx:chart's chart/plot fill only accepts solid color — not gradient
#     (unlike regular cChart). Use a single hex like "F8FAFC" or "none".
#   NOTE 2: `colors=` palette is omitted — cx single-series charts paint only
#     the first palette entry. Let Excel's theme drive per-segment coloring.
officecli add "$FILE" "/2-Treemap & Sunburst" --type chart \
  --prop chartType=sunburst \
  --prop title="Market Share by Region" \
  --prop series1=Share:35,25,20,15,30,25,20,10,15 \
  --prop categories=North,South,East,West,Urban,Suburban,Rural,Online,Retail \
  --prop chartFill=F8FAFC \
  --prop plotFill=FFFFFF \
  --prop dataLabels=true \
  --prop x=14 --prop y=19 --prop width=13 --prop height=18

# ==========================================================================
# Sheet 3: Histogram & Box Whisker
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="3-Histogram & BoxWhisker"

# Chart 1: Histogram — auto-binning (Excel picks bin count)
# Features: chartType=histogram, no binning knobs → Excel auto-selects bins
officecli add "$FILE" "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=histogram \
  --prop title="Test Scores (auto bins)" \
  --prop series1=Scores:45,52,58,61,63,65,67,68,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,97,99 \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# Chart 2: Histogram — explicit binCount=5 with title glow
# Features: histogram binCount (explicit bin count), title.glow
officecli add "$FILE" "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=histogram \
  --prop title="Sales (binCount=5)" \
  --prop series1=Sales:120,135,148,155,162,170,175,183,191,200,210,220,235,250,265,280,295,310,340,380,420,480,550,620,700 \
  --prop binCount=5 \
  --prop title.glow=FFC000-6-50 \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# Chart 3: Histogram — explicit binSize=50 (fixed bin width) + label font
# Features: histogram binSize (explicit bin width — mutually exclusive with
#   binCount), dataLabels, labelFont
officecli add "$FILE" "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=histogram \
  --prop title="Sales (binSize=50)" \
  --prop series1=Sales:120,135,148,155,162,170,175,183,191,200,210,220,235,250,265,280,295,310,340,380,420,480,550,620,700 \
  --prop binSize=50 \
  --prop dataLabels=true \
  --prop labelFont=9:FFFFFF:true \
  --prop x=28 --prop y=0 --prop width=13 --prop height=18

# Chart 4: Histogram — overflow/underflow bins + intervalClosed=l
# Features: histogram underflowBin (cutoff for <N), overflowBin (cutoff for >N),
#   intervalClosed=l (bins are [a,b) — left-closed; default "r" is (a,b]),
#   legend=none
officecli add "$FILE" "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=histogram \
  --prop title="Response Time (outlier bins)" \
  --prop series1=ms:40,55,68,75,82,88,95,102,110,118,125,135,150,175,220,280,350 \
  --prop underflowBin=60 \
  --prop overflowBin=200 \
  --prop intervalClosed=l \
  --prop dataLabels=true \
  --prop legend=none \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# Chart 5: Box & Whisker — two teams, quartileMethod=exclusive
# Features: chartType=boxWhisker, two-series comparison, quartileMethod=exclusive,
#   legend=bottom, outlier detection (built-in)
officecli add "$FILE" "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Response Time by Team (ms)" \
  --prop "series1=TeamA:42,55,61,68,72,75,78,81,85,88,92,97,105,120" \
  --prop "series2=TeamB:30,38,45,52,58,62,65,68,71,74,78,85,92,110" \
  --prop quartileMethod=exclusive \
  --prop legend=bottom \
  --prop x=14 --prop y=19 --prop width=13 --prop height=18

# Chart 6: Box & Whisker — three departments, quartileMethod=inclusive + glow
# Features: boxWhisker three-series, quartileMethod=inclusive (different quartile
#   formula from exclusive), title.glow, mean markers (default on)
officecli add "$FILE" "/3-Histogram & BoxWhisker" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Salary Distribution (\$k)" \
  --prop "series1=Engineering:85,92,95,98,102,105,108,112,118,125,135,150,180" \
  --prop "series2=Marketing:60,65,68,72,75,78,80,83,88,92,98,110" \
  --prop "series3=Sales:55,62,68,75,82,90,98,105,115,125,140,160,190" \
  --prop quartileMethod=inclusive \
  --prop title.glow=00D2FF-6-60 \
  --prop legend=bottom \
  --prop x=28 --prop y=19 --prop width=13 --prop height=18

# ==========================================================================
# Sheet 4: Pareto
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="4-Pareto"

# Chart 1: Pareto — defect analysis, raw counts auto-sorted + cumul% overlay
# Features: chartType=pareto (2-series under the hood — clusteredColumn bars +
#   paretoLine cumulative %), automatic descending sort, cumulative % computed
#   server-side, dataLabels on both series.
officecli add "$FILE" "/4-Pareto" --type chart \
  --prop chartType=pareto \
  --prop title="Defect Pareto" \
  --prop series1=Count:45,30,10,8,5,2 \
  --prop categories=Scratches,Dents,Cracks,Chips,Stains,Other \
  --prop dataLabels=true \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# Chart 2: Pareto — root cause analysis, 10 categories, out-of-order input
# Features: pareto with unsorted input values (12, 87, 5, ...) — officecli
#   re-sorts by value desc and re-aligns categories. title.glow + legend=bottom.
officecli add "$FILE" "/4-Pareto" --type chart \
  --prop chartType=pareto \
  --prop title="Root Cause Pareto" \
  --prop series1=Tickets:12,87,5,45,3,120,22,67,8,31 \
  --prop categories=Network,Auth,DB,Cache,UI,Config,Deploy,Monitor,Queue,Storage \
  --prop title.glow=FFC000-6-50 \
  --prop legend=bottom \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# ==========================================================================
# Sheet 5: Chart Meta
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="5-Chart Meta"

# Chart 1: anchor (cell-range placement), preset (named style bundle)
# Features: anchor="A1:M20" (position chart at exact cell-range instead of
#   x/y/width/height), preset=corporate (named style bundle: colors, fonts,
#   fill, border in one shot; values: minimal, dark, corporate, magazine,
#   dashboard, colorful, monochrome)
officecli add "$FILE" "/5-Chart Meta" --type chart \
  --prop chartType=column \
  --prop title="anchor + preset=corporate" \
  --prop series1=Revenue:120,145,132,160 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop anchor=A1:M20 \
  --prop preset=corporate

# Chart 2: autotitledeleted, plotvisonly
# Features: autotitledeleted=true (suppress the auto "Chart Title" placeholder),
#   plotvisonly=true (skip plotting hidden rows/columns)
officecli add "$FILE" "/5-Chart Meta" --type chart \
  --prop chartType=bar \
  --prop series1=Sales:80,95,88,110 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=0 --prop y=22 --prop width=12 --prop height=18 \
  --prop autotitledeleted=true \
  --prop plotvisonly=true

# Chart 3: preset variants — minimal
# Features: preset=minimal (strip: removes gridlines, legend, border, most
#   styling; exposes the data with minimal chrome)
officecli add "$FILE" "/5-Chart Meta" --type chart \
  --prop chartType=line \
  --prop title="preset=minimal" \
  --prop series1=A:10,20,15,25 \
  --prop series2=B:8,14,12,20 \
  --prop categories=W1,W2,W3,W4 \
  --prop x=13 --prop y=0 --prop width=12 --prop height=18 \
  --prop preset=minimal

# Chart 4: preset=dark
# Features: preset=dark (dark background, light-colored series and text)
officecli add "$FILE" "/5-Chart Meta" --type chart \
  --prop chartType=column \
  --prop title="preset=dark" \
  --prop series1=Sales:45,60,55,80 \
  --prop categories=Q1,Q2,Q3,Q4 \
  --prop x=13 --prop y=22 --prop width=12 --prop height=18 \
  --prop preset=dark

# Remove blank default Sheet1 (all data is inline)
officecli remove "$FILE" /Sheet1

officecli close "$FILE"
officecli validate "$FILE"
echo "Generated: $FILE"
