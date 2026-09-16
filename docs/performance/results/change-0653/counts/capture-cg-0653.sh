#!/usr/bin/env bash
# change 0653: callgrind capture for one leg (before|after).
# Codec isolation pairs plus whole-process eager / source-backed reads on the
# real producer fixture and change 0587's marker-stripped control.
set -u
LEG="$1"; BIN="$2"; S="$3"; CPU="${4:-8}"; WHICH="${5:-all}"
OUT="$S/cg-$LEG"; mkdir -p "$OUT"
SYMS='process_markup_compatibility|write_start|codec::start|codec::esc|BoundedOutput|for_each_hoisted|for_each_effective|__rust_realloc|__rust_alloc|from_utf8|memmem|find_bytes|read_event_impl|PROGRAM TOTALS'
run() {
  local name="$1"; shift
  taskset -c "$CPU" valgrind --tool=callgrind --callgrind-out-file="$OUT/$name.out" \
    --compress-strings=no --compress-pos=no "$BIN" "$@" \
    > "$OUT/$name.stdout" 2> "$OUT/$name.stderr"
  callgrind_annotate --inclusive=yes "$OUT/$name.out" 2>/dev/null | grep -E "$SYMS" | head -n 40 > "$OUT/$name.inclusive.txt"
  callgrind_annotate --inclusive=no  "$OUT/$name.out" 2>/dev/null | head -n 45 > "$OUT/$name.self.txt"
  rm -f "$OUT/$name.out"
}
if [ "$WHICH" = all ] || [ "$WHICH" = mce ]; then
  run mce-real-r1      mce "$S/real-sheet1.xml" 1
  run mce-real-r6      mce "$S/real-sheet1.xml" 6
  run mce-control-r1   mce "$S/control-sheet1.xml" 1
  run mce-control-r6   mce "$S/control-sheet1.xml" 6
fi
if [ "$WHICH" = all ] || [ "$WHICH" = public ]; then
  run eager-real     eager  "$S/real.xlsx"    H680
  run eager-control  eager  "$S/control.xlsx" H680
  run source-real    source "$S/real.xlsx"    H680
  run source-control source "$S/control.xlsx" H680
fi
echo "capture $LEG ($WHICH) done"
