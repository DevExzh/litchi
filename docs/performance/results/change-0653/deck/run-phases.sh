#!/usr/bin/env bash
# change 0653: change 0649's real-deck edit phases, both legs, ordered A1 B1 B2
# A2 per round with a repeated A/A block in the same window. One pinned process
# per block; the per-repeat spread of one leg is the floor.
set -u
M="$1"; B="$2"; A="$3"; REAL="$4"; CONTROL="$5"; CPU="${6:-8}"; N="${7:-30}"
OUT="$M/deck"; mkdir -p "$OUT"
for deck in real control; do
  case "$deck" in real) SRC="$REAL";; control) SRC="$CONTROL";; esac
  for round in 1 2; do
    taskset -c "$CPU" "$B" phases "$SRC" "$N" > "$OUT/$deck.before.$((round*2-1)).tsv"
    taskset -c "$CPU" "$A" phases "$SRC" "$N" > "$OUT/$deck.after.$((round*2-1)).tsv"
    taskset -c "$CPU" "$A" phases "$SRC" "$N" > "$OUT/$deck.after.$((round*2)).tsv"
    taskset -c "$CPU" "$B" phases "$SRC" "$N" > "$OUT/$deck.before.$((round*2)).tsv"
  done
  # A/A floor: two more before blocks in the same window.
  taskset -c "$CPU" "$B" phases "$SRC" "$N" > "$OUT/$deck.floor.1.tsv"
  taskset -c "$CPU" "$B" phases "$SRC" "$N" > "$OUT/$deck.floor.2.tsv"
done
echo "deck phases done"
