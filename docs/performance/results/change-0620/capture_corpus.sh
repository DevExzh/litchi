#!/usr/bin/env bash
# capture_corpus.sh <before-binary> <after-binary> <out-dir>
#
# The correctness oracle: every `.xls` in the repository, through all five
# `cell_values` publication paths, on both legs, three times each. Each line is
# either the exact typed refusal text or the SHA-256 of the complete published
# artifact with its diagnostics and patch lengths.
#
# Three runs per leg are not redundancy. The generic `Transaction::commit` path
# publishes a CFB whose storage directory entries are ordered by a `HashSet`
# iteration, so two fixtures publish different bytes on each process; repeating
# the run is what separates that pre-existing nondeterminism from a real
# before/after difference.
set -euo pipefail
BEFORE=$(readlink -f "$1")
AFTER=$(readlink -f "$2")
mkdir -p "$3"; OUT=$(readlink -f "$3")
REPO=${REPO:-/home/zhuhe/code/litchi}
cd "$REPO"
mapfile -t FIXTURES < <(find test-data -iname '*.xls' -type f | sort)
echo "${#FIXTURES[@]} fixtures"
for round in 1 2 3; do
  "$BEFORE" "${FIXTURES[@]}" > "$OUT/corpus-before-r$round.jsonl"
  "$AFTER"  "${FIXTURES[@]}" > "$OUT/corpus-after-r$round.jsonl"
  echo "ran round $round"
done
