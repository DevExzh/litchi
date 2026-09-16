#!/usr/bin/env bash
# perf stat isolation pairs: profile N and N+M operations in one process,
# difference the totals and divide by M. Pinned to CPU 19.
set -u
BIN="$1"; MODE="$2"; FIXTURE="$3"; N1="$4"; N2="$5"
run() {
  taskset -c 14 perf stat -r 3 -x, -e cycles,instructions \
    -- "$BIN" profile "$MODE" "$FIXTURE" 3 "$1" 2>&1 \
    | awk -F, '$3=="cycles"{c=$1} $3=="instructions"{i=$1} END{print c" "i}'
}
A=$(run "$N1"); B=$(run "$N2")
python3 - "$MODE" "$FIXTURE" "$N1" "$N2" $A $B <<'PY'
import os, sys
mode, fixture, n1, n2, c1, i1, c2, i2 = sys.argv[1:9]
m = int(n2) - int(n1)
cyc = (float(c2) - float(c1)) / m
ins = (float(i2) - float(i1)) / m
size = os.path.getsize(fixture)
print(f"{mode}\t{os.path.basename(fixture)}\t{size}\t{cyc:.0f}\t{ins:.0f}\t{cyc/size:.4f}\t{ins/size:.4f}")
PY
