#!/bin/bash
# Box-Whisker Chart Showcase — generates charts-boxwhisker.xlsx exercising the
# full xlsx boxWhisker chart property surface across 8 charts on 2 sheets.
# CLI twin of charts-boxwhisker.py (officecli Python SDK).
# Usage: ./charts-boxwhisker.sh
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-boxwhisker.xlsx"
rm -f "$FILE"

officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Sheet 1: Basics & Quartile Methods
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="1-Basics & Quartile"

# --------------------------------------------------------------------------
# Chart 1: Basic single-series with exclusive quartile and data labels
# Features: single series, quartileMethod=exclusive, dataLabels
# --------------------------------------------------------------------------
officecli add "$FILE" "/1-Basics & Quartile" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Test Score Distribution" \
  --prop "series1=Scores:45,52,58,61,63,65,67,68,70,72,75,78,82,85,90,95,99" \
  --prop quartileMethod=exclusive \
  --prop dataLabels=true \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# --------------------------------------------------------------------------
# Chart 2: Multi-series with inclusive quartile, legend at bottom
# Features: 3 series, quartileMethod=inclusive, legend=bottom
# --------------------------------------------------------------------------
officecli add "$FILE" "/1-Basics & Quartile" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Salary by Department (\$k)" \
  --prop "series1=Engineering:85,92,95,98,102,105,108,112,118,125,135,150,180" \
  --prop "series2=Marketing:60,65,68,72,75,78,80,83,88,92,98,110" \
  --prop "series3=Sales:55,62,68,75,82,90,98,105,115,125,140,160,190" \
  --prop quartileMethod=inclusive \
  --prop legend=bottom \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# --------------------------------------------------------------------------
# Chart 3: Title styling — color, size, bold, font, shadow
# Features: title.color, title.size, title.bold, title.font, title.shadow
# --------------------------------------------------------------------------
officecli add "$FILE" "/1-Basics & Quartile" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Styled Title Demo" \
  --prop title.color=1B2838 \
  --prop title.size=20 \
  --prop title.bold=true \
  --prop title.font=Georgia \
  --prop title.shadow=000000-6-45-3-50 \
  --prop "series1=Data:18,22,25,28,30,32,35,38,40,42,45,55,62,78" \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# --------------------------------------------------------------------------
# Chart 4: Series colors — fill, colors (per-series), series.shadow
# Features: colors (per-series hex), series.shadow
# --------------------------------------------------------------------------
officecli add "$FILE" "/1-Basics & Quartile" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Custom Series Colors" \
  --prop "series1=GroupA:30,38,45,52,58,62,65,68,71,74,78,85,92" \
  --prop "series2=GroupB:20,28,35,40,48,55,60,66,70,80,88,95,110" \
  --prop colors=5B9BD5,ED7D31 \
  --prop series.shadow=000000-6-45-3-35 \
  --prop x=14 --prop y=19 --prop width=13 --prop height=18

# ==========================================================================
# Sheet 2: Axes & Styling
# ==========================================================================
officecli add "$FILE" / --type sheet --prop name="2-Axes & Styling"

# --------------------------------------------------------------------------
# Chart 5: Axis scaling + axis titles + axis title styling + axis font
# Features: axismin, axismax, majorunit, minorunit, xAxisTitle, yAxisTitle,
#   axisTitle.color/.size/.bold/.font, axisfont
# --------------------------------------------------------------------------
officecli add "$FILE" "/2-Axes & Styling" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Response Time (ms)" \
  --prop "series1=API:12,18,22,25,28,30,32,35,38,40,42,45,55,62,78,95,120" \
  --prop "series2=DB:5,8,10,12,14,16,18,20,22,25,28,32,38,45,60" \
  --prop axismin=0 --prop axismax=130 --prop majorunit=10 --prop minorunit=5 \
  --prop xAxisTitle=Service \
  --prop yAxisTitle="Latency (ms)" \
  --prop axisTitle.color=4A5568 \
  --prop axisTitle.size=12 \
  --prop axisTitle.bold=true \
  --prop axisTitle.font="Helvetica Neue" \
  --prop "axisfont=10:6B7280:Consolas" \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# --------------------------------------------------------------------------
# Chart 6: Axis visibility + axis lines + gridlines + xGridlines
# Features: cataxis.visible=false, valaxis.line, gridlines, gridlineColor,
#   xGridlines, xGridlineColor
# --------------------------------------------------------------------------
officecli add "$FILE" "/2-Axes & Styling" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Axis & Gridline Control" \
  --prop "series1=Temp:15,18,20,22,24,26,28,30,32,35,38,40,42" \
  --prop cataxis.visible=false \
  --prop "valaxis.line=334155:1.5" \
  --prop gridlines=true \
  --prop gridlineColor=E2E8F0 \
  --prop xGridlines=true \
  --prop xGridlineColor=F1F5F9 \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# --------------------------------------------------------------------------
# Chart 7: Plot/chart area fills, borders, gapWidth, tickLabels=false
# Features: fill (single color), gapWidth, tickLabels=false, gridlines=false,
#   plotareafill, plotarea.border, chartareafill, chartarea.border
# --------------------------------------------------------------------------
officecli add "$FILE" "/2-Axes & Styling" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Card Style" \
  --prop "series1=Weight:50,55,58,60,62,64,66,68,70,72,75,78,82,88,95" \
  --prop fill=6366F1 \
  --prop gapWidth=200 \
  --prop tickLabels=false \
  --prop gridlines=false \
  --prop plotareafill=F8FAFC \
  --prop "plotarea.border=E2E8F0:1" \
  --prop chartareafill=FFFFFF \
  --prop "chartarea.border=CBD5E1:0.75" \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# --------------------------------------------------------------------------
# Chart 8: Full presentation-grade — everything combined
# Features: ALL properties combined — title styling, multi-series colors,
#   series.shadow, axis scaling, axis titles + styling, axisfont, axisline,
#   gridlineColor, dataLabels + numfmt, legend + overlay + legendfont,
#   plot/chart area fill + border
# --------------------------------------------------------------------------
officecli add "$FILE" "/2-Axes & Styling" --type chart \
  --prop chartType=boxWhisker \
  --prop title="Server Latency Dashboard" \
  --prop title.color=0F172A \
  --prop title.size=18 \
  --prop title.bold=true \
  --prop title.font="Helvetica Neue" \
  --prop title.shadow=000000-4-45-2-40 \
  --prop "series1=US-East:8,12,15,18,20,22,24,26,28,30,35,42,55,70,95" \
  --prop "series2=EU-West:10,14,18,22,25,28,30,33,36,40,45,50,60,80" \
  --prop "series3=AP-South:15,20,25,30,35,38,42,45,48,52,58,65,75,90,120" \
  --prop quartileMethod=exclusive \
  --prop colors=3B82F6,10B981,F59E0B \
  --prop series.shadow=000000-4-45-2-30 \
  --prop axismin=0 --prop axismax=130 --prop majorunit=10 \
  --prop xAxisTitle=Region \
  --prop yAxisTitle="Latency (ms)" \
  --prop axisTitle.color=475569 \
  --prop axisTitle.size=11 \
  --prop axisTitle.bold=true \
  --prop axisTitle.font="Helvetica Neue" \
  --prop "axisfont=9:64748B:Helvetica Neue" \
  --prop "axisline=CBD5E1:1" \
  --prop gridlineColor=E2E8F0 \
  --prop dataLabels=true \
  --prop "datalabels.numfmt=0" \
  --prop legend=top \
  --prop legend.overlay=false \
  --prop "legendfont=10:475569:Helvetica Neue" \
  --prop plotareafill=F8FAFC \
  --prop "plotarea.border=E2E8F0:0.75" \
  --prop chartareafill=FFFFFF \
  --prop "chartarea.border=CBD5E1:0.75" \
  --prop x=14 --prop y=19 --prop width=16 --prop height=22

# Remove blank default Sheet1
officecli remove "$FILE" /Sheet1

officecli close "$FILE"

officecli validate "$FILE"
echo "Generated: $FILE"
