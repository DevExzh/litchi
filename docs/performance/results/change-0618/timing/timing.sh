#!/usr/bin/env bash
# Paired timing for change 0618. A = before leg (8fe9efa55), B = after leg.
# Legs run in A1 B1 B2 A2 order so the A/A pair brackets the B pair and both
# see the same window of host load.
set -euo pipefail
S="$1"; CPU=12
CFS=/home/zhuhe/code/litchi-worktrees/before-8fe9efa55/test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx
OPC_A=/home/zhuhe/code/litchi-worktrees/targets/0618-opc-before/release/probe0618-opc
OPC_B=/home/zhuhe/code/litchi-worktrees/targets/0618-opc-after/release/probe0618-opc
PPTX_A=/home/zhuhe/code/litchi-worktrees/targets/0618-p0607-before/release/probe0618-pptx
PPTX_B=/home/zhuhe/code/litchi-worktrees/targets/0618-p0607-after/release/probe0618-pptx

# --- preservation writer: per-sample publish nanoseconds -------------------
opc_leg() { # opc_leg <dir> <binary> <name> <scenario>
  local dir="$1" bin="$2" name="$3" scenario="$4"
  mkdir -p "$dir"
  taskset -c $CPU "$bin" time "$CFS" "$scenario" 200 2000 > "$dir/$name.txt" 2>/dev/null
}
for scenario in addrelall addrel; do
  d="$S/timing/opc-$scenario"
  opc_leg "$d" "$OPC_A" A1 "$scenario"
  opc_leg "$d" "$OPC_B" B1 "$scenario"
  opc_leg "$d" "$OPC_B" B2 "$scenario"
  opc_leg "$d" "$OPC_A" A2 "$scenario"
done

# --- streaming writer: mean nanoseconds per authored save ------------------
pptx_leg() { # pptx_leg <dir> <binary> <name> <slides> <iters> <samples>
  local dir="$1" bin="$2" name="$3" slides="$4" iters="$5" samples="$6"
  mkdir -p "$dir"
  : > "$dir/$name.txt"
  for _ in 1 2 3; do taskset -c $CPU "$bin" create "$slides" "$iters" >/dev/null; done
  for _ in $(seq 1 "$samples"); do
    taskset -c $CPU "$bin" create "$slides" "$iters" | awk -F'\t' '{print $1/$2}' >> "$dir/$name.txt"
  done
}
d="$S/timing/pptx-create50"
pptx_leg "$d" "$PPTX_A" A1 50 8 40
pptx_leg "$d" "$PPTX_B" B1 50 8 40
pptx_leg "$d" "$PPTX_B" B2 50 8 40
pptx_leg "$d" "$PPTX_A" A2 50 8 40
