#!/bin/bash
# Histogram Charts — Grand Showcase (officecli CLI twin of charts-histogram.py)
#
# The most thorough, most visually polished histogram demo officecli can
# produce. Every binning knob, every styling vocabulary, every canonical
# distribution shape, six design themes on one dataset, four font type
# specimens, and a cohesive production-grade ML dashboard — all driven by
# real copyable officecli CLI commands.
#
# Generates: charts-histogram.xlsx (6 sheets, 29 histograms)
#
#   0-Hero                 1 magazine-grade full-bleed hero poster chart
#   1-Binning Lab          6 charts — every binning knob, identical styling
#   2-Distribution Zoo     6 canonical real-world distribution shapes
#   3-Theme Gallery        6 design themes on the SAME dataset
#   4-Typography           4 font-family type specimens
#   5-ML Dashboard         6-chart "Production ML Model Report" dashboard
#
# Usage:
#   ./charts-histogram.sh

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/charts-histogram.xlsx"

# --------------------------------------------------------------------------
# Deterministic sample generators — same seed, same file every regeneration.
# bash has no Gaussian RNG, so the datasets are produced once by an inline
# python3 helper (identical seeds/maths to charts-histogram.py) and read into
# shell variables. Everything below is plain `officecli ... --prop k=v`.
# --------------------------------------------------------------------------
eval "$(python3 - <<'PY'
import random, math
def csv(v): return ",".join(str(x) for x in v)
def emit(name, vals): print(f'{name}="{csv(vals)}"')

random.seed(42); emit("BELL_CSV", sorted(round(random.gauss(75,12),1) for _ in range(200)))
random.seed(7);  emit("BIMODAL_CSV", sorted([round(random.gauss(55,6),1) for _ in range(80)]+[round(random.gauss(88,5),1) for _ in range(80)]))
random.seed(11); emit("LOGNORM_CSV", sorted(round(math.exp(random.gauss(3.2,0.55)),1) for _ in range(180)))
random.seed(23); emit("LEFT_CSV", sorted(round(75-math.exp(random.gauss(1.6,0.6)),1) for _ in range(140)))
random.seed(31); emit("UNIFORM_CSV", sorted(round(random.uniform(0,100),1) for _ in range(160)))
random.seed(47); emit("PARETO_CSV", sorted(round(random.paretovariate(1.6)*20,1) for _ in range(200)))
random.seed(101);emit("LATENCY_CSV", sorted(round(random.paretovariate(1.8)*15+10,1) for _ in range(250)))
random.seed(102);emit("CONFIDENCE_CSV", sorted(round(random.betavariate(6,2)*100,2) for _ in range(240)))
random.seed(103);emit("ERROR_MAG_CSV", sorted(round(abs(random.gauss(0,1.5)),3) for _ in range(180)))
random.seed(104);emit("TOKEN_CSV", sorted([max(1,round(random.gauss(180,40))) for _ in range(100)]+[max(1,round(random.gauss(520,90))) for _ in range(80)]))
random.seed(105);emit("GPU_CSV", sorted(round(min(99.0,max(30.0,random.gauss(82,8))),1) for _ in range(200)))
random.seed(106);emit("COST_CSV", sorted(round(math.exp(random.gauss(-3.2,0.9))*1000,3) for _ in range(220)))
PY
)"

rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"

# ==========================================================================
# Sheet 0: "0-Hero" — the full-bleed magazine hero poster
#
# A single giant chart using EVERY histogram knob at once: dark "Midnight
# Academia" palette, title.*, series.shadow + fill, axisline + axisfont +
# axisTitle.*, plotarea/chartarea fill + border, axismin/axismax/majorunit,
# gridlineColor, dataLabels + datalabels.numfmt, legend=top + legend.overlay
# + legendfont, intervalClosed=l + explicit binCount. The "representative
# sample" — if it renders correctly, the entire histogram pipeline is healthy.
# ==========================================================================
echo "--- 0-Hero ---"
officecli set "$FILE" /Sheet1 --prop name="0-Hero"
officecli add "$FILE" "/0-Hero" --type chart \
  --prop chartType=histogram \
  --prop title="The Shape of Data · 200-sample bell curve" \
  --prop title.color=F5F1E0 --prop title.size=22 --prop title.bold=true \
  --prop title.font="Helvetica Neue" \
  --prop title.shadow=000000-8-45-4-70 \
  --prop series1="Samples:$BELL_CSV" \
  --prop binCount=24 --prop intervalClosed=l \
  --prop fill=F0C96A --prop series.shadow=000000-8-45-4-60 \
  --prop axismin=0 --prop axismax=28 --prop majorunit=4 \
  --prop xAxisTitle="Score" --prop yAxisTitle="Frequency" \
  --prop axisTitle.color=C9B87A --prop axisTitle.size=13 \
  --prop axisTitle.bold=true --prop axisTitle.font="Helvetica Neue" \
  --prop axisfont="10:B8B090:Helvetica Neue" \
  --prop axisline="6A6448:1.5" \
  --prop gridlineColor=2F3544 \
  --prop plotareafill=1A1F2C --prop plotarea.border="3A3E4E:1.25" \
  --prop chartareafill=0B0F18 --prop chartarea.border="2A2E3E:1" \
  --prop dataLabels=true --prop datalabels.numfmt=0 \
  --prop legend=top --prop legend.overlay=false \
  --prop legendfont="11:D4C994:Helvetica Neue" \
  --prop x=0 --prop y=0 --prop width=27 --prop height=38

# ==========================================================================
# Sheet 1: "1-Binning Lab"
#
# Six histograms, SAME dataset (BELL), IDENTICAL typography / colors / frames
# — the ONLY thing that varies is the binning strategy. The "Rosetta stone":
# once you see how each binning knob reshapes the bars, you'll never be
# confused about which to use.
# ==========================================================================
echo "--- 1-Binning Lab ---"
officecli add "$FILE" / --type sheet --prop name="1-Binning Lab"

# Shared "clean lab" style — every chart wears the exact same outfit so the
# bin-shape difference is the only visible variable.
LAB=(
  --prop fill=4472C4
  --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true
  --prop title.font="Helvetica Neue"
  --prop xAxisTitle="Score" --prop yAxisTitle="Count"
  --prop axisTitle.color=6B7280 --prop axisTitle.size=10
  --prop axisTitle.font="Helvetica Neue"
  --prop axisfont="9:6B7280:Helvetica Neue"
  --prop gridlineColor=F0F0F0
  --prop plotareafill=FFFFFF --prop plotarea.border="E5E7EB:0.75"
  --prop chartareafill=F9FAFB --prop chartarea.border="E5E7EB:0.75"
  --prop axisline="9CA3AF:0.75"
)

# no binCount, no binSize — Excel picks the bin count automatically.
officecli add "$FILE" "/1-Binning Lab" --type chart \
  --prop chartType=histogram --prop title="1 · Auto-binning (Excel default)" \
  --prop series1="Samples:$BELL_CSV" "${LAB[@]}" \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# binCount=8 — coarse. Fewer, wider bars. Good for "what's the mode?"
officecli add "$FILE" "/1-Binning Lab" --type chart \
  --prop chartType=histogram --prop title="2 · binCount=8 (coarse)" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=8 "${LAB[@]}" \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# binCount=32 — fine. Many narrow bars. Good for "is it really Gaussian?"
officecli add "$FILE" "/1-Binning Lab" --type chart \
  --prop chartType=histogram --prop title="3 · binCount=32 (fine)" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=32 "${LAB[@]}" \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# binSize=5 — fixed bin width. Human-friendly bin boundaries regardless of range.
officecli add "$FILE" "/1-Binning Lab" --type chart \
  --prop chartType=histogram --prop title="4 · binSize=5 (fixed-width bins)" \
  --prop series1="Samples:$BELL_CSV" --prop binSize=5 "${LAB[@]}" \
  --prop x=14 --prop y=19 --prop width=13 --prop height=18

# underflowBin=55 + overflowBin=95 — outlier fencing. Everything below 55 or
# above 95 collapses into a single <55 / >95 bar.
officecli add "$FILE" "/1-Binning Lab" --type chart \
  --prop chartType=histogram --prop title="5 · underflow=55 · overflow=95 (fencing)" \
  --prop series1="Samples:$BELL_CSV" --prop binSize=5 --prop underflowBin=55 --prop overflowBin=95 "${LAB[@]}" \
  --prop x=0 --prop y=38 --prop width=13 --prop height=18

# intervalClosed=l (half-open [a,b)) + gapWidth=30 — left-closed variant AND
# bars pushed apart. Useful when values lie exactly on a bin boundary.
officecli add "$FILE" "/1-Binning Lab" --type chart \
  --prop chartType=histogram --prop title="6 · [a,b) intervals + gapWidth=30" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=16 --prop intervalClosed=l --prop gapWidth=30 "${LAB[@]}" \
  --prop x=14 --prop y=38 --prop width=13 --prop height=18

# ==========================================================================
# Sheet 2: "2-Distribution Zoo"
#
# A cohesive 2x3 gallery of the canonical distribution shapes you'll see in
# production data. Same typography + frames; only the fill color, data, and
# binning strategy change.
# ==========================================================================
echo "--- 2-Distribution Zoo ---"
officecli add "$FILE" / --type sheet --prop name="2-Distribution Zoo"

ZOO=(
  --prop title.color=1F2937 --prop title.size=13 --prop title.bold=true
  --prop title.font="Helvetica Neue"
  --prop axisTitle.color=6B7280 --prop axisTitle.size=10
  --prop axisTitle.font="Helvetica Neue"
  --prop axisfont="9:6B7280:Helvetica Neue"
  --prop gridlineColor=EFEFEF
  --prop plotareafill=FFFFFF --prop plotarea.border="E5E7EB:0.75"
  --prop chartareafill=F9FAFB --prop chartarea.border="E5E7EB:0.75"
  --prop axisline="9CA3AF:0.75"
)

# classic bell curve reference, binCount=18, midnight blue fill.
officecli add "$FILE" "/2-Distribution Zoo" --type chart \
  --prop chartType=histogram --prop title="Normal · bell curve (reference)" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=2F5597 \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" "${ZOO[@]}" \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# bimodal — two hidden populations. Narrow bins reveal the split.
officecli add "$FILE" "/2-Distribution Zoo" --type chart \
  --prop chartType=histogram --prop title="Bimodal · two hidden cohorts" \
  --prop series1="Score:$BIMODAL_CSV" --prop binCount=22 --prop fill=ED7D31 \
  --prop xAxisTitle="Test score" --prop yAxisTitle="Students" "${ZOO[@]}" \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# right-skewed log-normal. Mean >> median, long tail to the right.
officecli add "$FILE" "/2-Distribution Zoo" --type chart \
  --prop chartType=histogram --prop title="Right-skewed · log-normal (income)" \
  --prop series1="Income:$LOGNORM_CSV" --prop binCount=20 --prop fill=70AD47 \
  --prop xAxisTitle="Monthly income (\$k)" --prop yAxisTitle="People" "${ZOO[@]}" \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# left-skewed — retirement ages cluster high, tail stretches left.
officecli add "$FILE" "/2-Distribution Zoo" --type chart \
  --prop chartType=histogram --prop title="Left-skewed · retirement ages" \
  --prop series1="Age:$LEFT_CSV" --prop binCount=18 --prop fill=7030A0 \
  --prop xAxisTitle="Age at retirement" --prop yAxisTitle="Retirees" "${ZOO[@]}" \
  --prop x=14 --prop y=19 --prop width=13 --prop height=18

# uniform — every value equally likely. binSize emphasizes the "flat floor".
officecli add "$FILE" "/2-Distribution Zoo" --type chart \
  --prop chartType=histogram --prop title="Uniform · flat floor" \
  --prop series1="Draws:$UNIFORM_CSV" --prop binSize=10 --prop fill=00B0F0 \
  --prop xAxisTitle="Random draw (0-100)" --prop yAxisTitle="Count" "${ZOO[@]}" \
  --prop x=0 --prop y=38 --prop width=13 --prop height=18

# heavy-tailed Pareto + overflowBin. Fences the catastrophic tail so the
# interesting bulk of the distribution stays readable.
officecli add "$FILE" "/2-Distribution Zoo" --type chart \
  --prop chartType=histogram --prop title="Heavy-tailed · Pareto (overflow=250)" \
  --prop series1="Latency:$PARETO_CSV" --prop binSize=20 --prop overflowBin=250 --prop fill=C00000 \
  --prop xAxisTitle="Latency (ms)" --prop yAxisTitle="Requests" "${ZOO[@]}" \
  --prop x=14 --prop y=38 --prop width=13 --prop height=18

# ==========================================================================
# Sheet 3: "3-Theme Gallery"
#
# Six complete design themes applied to the SAME bell-curve dataset. Each
# theme is a coordinated palette chosen to read as one coherent mood.
# ==========================================================================
echo "--- 3-Theme Gallery ---"
officecli add "$FILE" / --type sheet --prop name="3-Theme Gallery"

# Theme 1 · Midnight Academia (dark plot area, gold bars, shadows)
officecli add "$FILE" "/3-Theme Gallery" --type chart \
  --prop chartType=histogram --prop title="Midnight Academia" \
  --prop title.color=F5F1E0 --prop title.size=14 --prop title.bold=true \
  --prop title.font="Georgia" --prop title.shadow=000000-6-45-3-70 \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=F0C96A \
  --prop series.shadow=000000-6-45-3-55 \
  --prop plotareafill=1A1F2C --prop plotarea.border="3A3E4E:1" \
  --prop chartareafill=0B0F18 --prop chartarea.border="2A2E3E:0.75" \
  --prop gridlineColor=2F3544 \
  --prop axisfont="9:B8B090:Georgia" \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=C9B87A --prop axisTitle.size=10 \
  --prop axisTitle.font="Georgia" --prop axisline="5A5848:1" \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# Theme 2 · Sunset Terracotta (warm cream + coral, serif)
officecli add "$FILE" "/3-Theme Gallery" --type chart \
  --prop chartType=histogram --prop title="Sunset Terracotta" \
  --prop title.color=3F2818 --prop title.size=14 --prop title.bold=true \
  --prop title.font="Georgia" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=E85D4A \
  --prop plotareafill=FFF5E8 --prop plotarea.border="F0D8B0:1" \
  --prop chartareafill=FFE6C7 --prop chartarea.border="E6BC88:1" \
  --prop gridlineColor=F5C98A \
  --prop axisfont="9:6B4A2A:Georgia" \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=A8522C --prop axisTitle.size=10 \
  --prop axisTitle.font="Georgia" --prop axisline="C08050:1" \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# Theme 3 · Forest Parchment (beige + forest green, serif)
officecli add "$FILE" "/3-Theme Gallery" --type chart \
  --prop chartType=histogram --prop title="Forest Parchment" \
  --prop title.color=1F3A1F --prop title.size=14 --prop title.bold=true \
  --prop title.font="Georgia" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=2F5D3A \
  --prop plotareafill=F3EDD8 --prop plotarea.border="C8B890:1" \
  --prop chartareafill=EADFBE --prop chartarea.border="A89858:1" \
  --prop gridlineColor=C0B888 \
  --prop axisfont="9:4A5A3A:Georgia" \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=3F5A2F --prop axisTitle.size=10 \
  --prop axisTitle.font="Georgia" --prop axisline="6A7A4A:1" \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# Theme 4 · Editorial Mono (pure grayscale, sans)
officecli add "$FILE" "/3-Theme Gallery" --type chart \
  --prop chartType=histogram --prop title="Editorial Mono" \
  --prop title.color=111111 --prop title.size=14 --prop title.bold=true \
  --prop title.font="Helvetica Neue" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=2A2A2A \
  --prop plotareafill=FFFFFF --prop plotarea.border="CCCCCC:0.75" \
  --prop chartareafill=FAFAFA --prop chartarea.border="E0E0E0:0.75" \
  --prop gridlineColor=EEEEEE \
  --prop axisfont="9:555555:Helvetica Neue" \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=333333 --prop axisTitle.size=10 \
  --prop axisTitle.font="Helvetica Neue" --prop axisline="888888:1" \
  --prop x=14 --prop y=19 --prop width=13 --prop height=18

# Theme 5 · Neon Terminal (black + electric cyan, mono)
officecli add "$FILE" "/3-Theme Gallery" --type chart \
  --prop chartType=histogram --prop title="Neon Terminal" \
  --prop title.color=00F0C8 --prop title.size=14 --prop title.bold=true \
  --prop title.font="Courier New" --prop title.shadow=00F0C8-6-45-0-40 \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=00F0C8 \
  --prop series.shadow=00F0C8-8-45-0-45 \
  --prop plotareafill=0A0A14 --prop plotarea.border="1F2F3F:1" \
  --prop chartareafill=000008 --prop chartarea.border="1F1F2F:1" \
  --prop gridlineColor=1A2A3A \
  --prop axisfont="9:00D0E8:Courier New" \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=00D0E8 --prop axisTitle.size=10 \
  --prop axisTitle.font="Courier New" --prop axisline="00707F:1" \
  --prop x=0 --prop y=38 --prop width=13 --prop height=18

# Theme 6 · Pastel Bloom (lavender cream + rose, sans)
officecli add "$FILE" "/3-Theme Gallery" --type chart \
  --prop chartType=histogram --prop title="Pastel Bloom" \
  --prop title.color=5A3C4A --prop title.size=14 --prop title.bold=true \
  --prop title.font="Helvetica Neue" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=F5A7C8 \
  --prop plotareafill=FDF4F8 --prop plotarea.border="F0D0E0:1" \
  --prop chartareafill=FAEDF2 --prop chartarea.border="F0C0D8:1" \
  --prop gridlineColor=F5D8E5 \
  --prop axisfont="9:8A6878:Helvetica Neue" \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=A04C6A --prop axisTitle.size=10 \
  --prop axisTitle.font="Helvetica Neue" --prop axisline="C888A0:1" \
  --prop x=14 --prop y=38 --prop width=13 --prop height=18

# ==========================================================================
# Sheet 4: "4-Typography"
#
# Four font-family "type specimens". Same data, same geometry, same colors —
# only the font varies. Helvetica is corporate, Georgia is editorial, Courier
# is data, Verdana is approachable.
# ==========================================================================
echo "--- 4-Typography ---"
officecli add "$FILE" / --type sheet --prop name="4-Typography"

# Specimen 1 · Helvetica Neue (modern sans — dashboards, corporate reports)
officecli add "$FILE" "/4-Typography" --type chart \
  --prop chartType=histogram --prop title="Helvetica Neue · modern sans" \
  --prop title.color=1F2937 --prop title.size=16 --prop title.bold=true \
  --prop title.font="Helvetica Neue" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=4472C4 \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=4472C4 --prop axisTitle.size=11 \
  --prop axisTitle.font="Helvetica Neue" \
  --prop axisfont="10:6B7280:Helvetica Neue" \
  --prop gridlineColor=EEEEEE \
  --prop plotareafill=FFFFFF --prop plotarea.border="E5E7EB:0.75" \
  --prop chartareafill=F9FAFB --prop chartarea.border="E5E7EB:0.75" \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# Specimen 2 · Georgia (editorial serif — magazines, long-form reports)
officecli add "$FILE" "/4-Typography" --type chart \
  --prop chartType=histogram --prop title="Georgia · editorial serif" \
  --prop title.color=3F2818 --prop title.size=16 --prop title.bold=true \
  --prop title.font="Georgia" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=A8522C \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=A8522C --prop axisTitle.size=11 \
  --prop axisTitle.font="Georgia" \
  --prop axisfont="10:6B4A2A:Georgia" \
  --prop gridlineColor=F0E8D8 \
  --prop plotareafill=FFFBF3 --prop plotarea.border="E8D8B8:0.75" \
  --prop chartareafill=FDF6E8 --prop chartarea.border="E8D8B8:0.75" \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# Specimen 3 · Courier New (monospace — data, telemetry, engineering)
officecli add "$FILE" "/4-Typography" --type chart \
  --prop chartType=histogram --prop title="Courier New · data mono" \
  --prop title.color=1A3A1A --prop title.size=16 --prop title.bold=true \
  --prop title.font="Courier New" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=2F8F4F \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=2F8F4F --prop axisTitle.size=11 \
  --prop axisTitle.font="Courier New" \
  --prop axisfont="10:3A5A3A:Courier New" \
  --prop gridlineColor=E0EDE0 \
  --prop plotareafill=F7FBF7 --prop plotarea.border="C8DCC8:0.75" \
  --prop chartareafill=F0F7F0 --prop chartarea.border="C8DCC8:0.75" \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# Specimen 4 · Verdana (friendly sans — onboarding, public-facing UI)
officecli add "$FILE" "/4-Typography" --type chart \
  --prop chartType=histogram --prop title="Verdana · friendly sans" \
  --prop title.color=4A2B6A --prop title.size=16 --prop title.bold=true \
  --prop title.font="Verdana" \
  --prop series1="Samples:$BELL_CSV" --prop binCount=18 --prop fill=8E4DBB \
  --prop xAxisTitle="Score" --prop yAxisTitle="Count" \
  --prop axisTitle.color=8E4DBB --prop axisTitle.size=11 \
  --prop axisTitle.font="Verdana" \
  --prop axisfont="10:6B4A8A:Verdana" \
  --prop gridlineColor=ECE0F4 \
  --prop plotareafill=FCF7FF --prop plotarea.border="D8C4E8:0.75" \
  --prop chartareafill=F6EDFA --prop chartarea.border="D8C4E8:0.75" \
  --prop x=14 --prop y=19 --prop width=13 --prop height=18

# ==========================================================================
# Sheet 5: "5-ML Dashboard"
#
# A cohesive six-chart "Production ML Model Report". Every chart wears the
# same corporate dashboard uniform but each shows a different slice of the
# model's behavior, with a different color + binning strategy.
#
#   Row 1:  Inference latency (ms)   |  Prediction confidence (%)
#   Row 2:  |Residual| (logit)       |  Token length (chars)
#   Row 3:  GPU utilization (%)      |  Cost per request ($ × 0.001)
# ==========================================================================
echo "--- 5-ML Dashboard ---"
officecli add "$FILE" / --type sheet --prop name="5-ML Dashboard"

DASH=(
  --prop title.color=1F2937 --prop title.size=12 --prop title.bold=true
  --prop title.font="Helvetica Neue"
  --prop axisTitle.color=6B7280 --prop axisTitle.size=9
  --prop axisTitle.font="Helvetica Neue"
  --prop axisfont="8:6B7280:Helvetica Neue"
  --prop gridlineColor=F0F0F0
  --prop plotareafill=FFFFFF --prop plotarea.border="E5E7EB:0.75"
  --prop chartareafill=F9FAFB --prop chartarea.border="E5E7EB:0.75"
  --prop axisline="9CA3AF:0.75"
  --prop dataLabels=false
)

# 1 · Inference Latency — heavy-tail, overflow-fenced, red for "watch this"
officecli add "$FILE" "/5-ML Dashboard" --type chart \
  --prop chartType=histogram --prop title="Inference Latency · p50-p99 (ms)" \
  --prop series1="Latency:$LATENCY_CSV" --prop binSize=25 --prop overflowBin=300 --prop fill=EF4444 \
  --prop series.shadow=EF4444-4-45-2-25 \
  --prop xAxisTitle="Latency (ms)" --prop yAxisTitle="Requests" "${DASH[@]}" \
  --prop x=0 --prop y=0 --prop width=13 --prop height=18

# 2 · Prediction Confidence — beta-like, axismin/max locked to 0..100
officecli add "$FILE" "/5-ML Dashboard" --type chart \
  --prop chartType=histogram --prop title="Prediction Confidence" \
  --prop series1="Confidence:$CONFIDENCE_CSV" --prop binSize=5 --prop fill=10B981 \
  --prop axismin=0 --prop majorunit=50 \
  --prop xAxisTitle="Softmax confidence (%)" --prop yAxisTitle="Samples" "${DASH[@]}" \
  --prop x=14 --prop y=0 --prop width=13 --prop height=18

# 3 · Residual Magnitude — half-normal, intervalClosed=l so bin=0 catches zeros
officecli add "$FILE" "/5-ML Dashboard" --type chart \
  --prop chartType=histogram --prop title="|Residual| · model calibration" \
  --prop series1="Residual:$ERROR_MAG_CSV" --prop binSize=0.25 --prop intervalClosed=l --prop fill=F59E0B \
  --prop xAxisTitle="|y - ŷ| (logit)" --prop yAxisTitle="Samples" "${DASH[@]}" \
  --prop x=0 --prop y=19 --prop width=13 --prop height=18

# 4 · Token Length — bimodal (short prompts vs long prompts)
officecli add "$FILE" "/5-ML Dashboard" --type chart \
  --prop chartType=histogram --prop title="Token Length · short vs long prompts" \
  --prop series1="Tokens:$TOKEN_CSV" --prop binCount=24 --prop fill=6366F1 \
  --prop xAxisTitle="Tokens" --prop yAxisTitle="Requests" "${DASH[@]}" \
  --prop x=14 --prop y=19 --prop width=13 --prop height=18

# 5 · GPU Utilization — locked axis range so dashboard charts share scale
officecli add "$FILE" "/5-ML Dashboard" --type chart \
  --prop chartType=histogram --prop title="GPU Utilization" \
  --prop series1="GPU:$GPU_CSV" --prop binSize=5 --prop fill=8B5CF6 \
  --prop axismin=0 --prop axismax=50 --prop majorunit=10 \
  --prop xAxisTitle="Utilization (%)" --prop yAxisTitle="Samples" "${DASH[@]}" \
  --prop x=0 --prop y=38 --prop width=13 --prop height=18

# 6 · Cost per Request — log-normal, overflow-fenced, data labels with numfmt.
# DASH carries dataLabels=false; this chart overrides it to true (last wins).
officecli add "$FILE" "/5-ML Dashboard" --type chart \
  --prop chartType=histogram --prop title="Cost per Request (\$ × 0.001)" \
  --prop series1="Cost:$COST_CSV" --prop binSize=5 --prop overflowBin=120 --prop fill=EC4899 \
  --prop xAxisTitle="Cost (m\$)" --prop yAxisTitle="Requests" "${DASH[@]}" \
  --prop dataLabels=true --prop datalabels.numfmt=0 \
  --prop x=14 --prop y=38 --prop width=13 --prop height=18

officecli close "$FILE"
officecli validate "$FILE"

echo "Done! Generated: $FILE"
echo "  6 sheets, 29 histograms total"
