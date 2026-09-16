#!/usr/bin/env bash
# Change 0658: callgrind isolation pairs for the source-backed one-cell read.
# Usage: run_callgrind.sh <leg> <binary>
# Profiles N=1 and N=11 repeats of the operation over fresh workbooks and
# differences the totals (M=10) to isolate one read. Raw callgrind outputs are
# deleted after the annotated tables are extracted.
set -u
LEG="$1"; BIN="$2"
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0658
cd "$S/out" || exit 1
SYMS='process_markup_compatibility|Parser>::parse|x14ac::capture|selected::scan|scan_range|Processor>::start|clone_bounded_name_part|read_event_impl|memmem|find_bytes|styles::parse|PROGRAM TOTALS|store|prefix|invoke_active|Processor>::end'
for v in control real; do
  for n in 1 11; do
    out="cg-$LEG-$v-$n.out"
    taskset -c 13 valgrind --tool=callgrind --callgrind-out-file="$out" \
      --compress-strings=no --compress-pos=no \
      "$BIN" ir "$S/fixtures/$v.xlsx" H680 "$n" >/dev/null 2>"cg-$LEG-$v-$n.stderr"
    callgrind_annotate --threshold=99.99 --inclusive=yes "$out" 2>/dev/null | grep -E "$SYMS" | head -n 60 > "cg-$LEG-$v-$n.inclusive.txt"
    callgrind_annotate --inclusive=no  "$out" 2>/dev/null | head -n 40 > "cg-$LEG-$v-$n.self.txt"
    rm -f "$out"
  done
done
echo "$LEG done"
