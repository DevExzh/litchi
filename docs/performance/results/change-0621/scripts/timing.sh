#!/bin/bash
# Paired timing, A1 B1 B2 A2 plus an A/A control, one fresh child per leg.
BEFORE_BIN="$1"; BEFORE_ROOT="$2"; AFTER_BIN="$3"; AFTER_ROOT="$4"; OUT="$5"
mkdir -p "$OUT"
leg() { # tag binary root mode op fixture samples
  local tag="$1" bin="$2" root="$3" m="$4" op="$5" f="$6" n="$7"
  taskset -c 16 setarch "$(uname -m)" -R "$bin" --input "$root/$f" --mode "$m" \
    --operation "$op" --warmups 3 --samples "$n" > "$OUT/$tag.json" 2>"$OUT/$tag.err"
}
case_run() { # label mode op fixture samples
  local lab="$1" m="$2" op="$3" f="$4" n="$5"
  leg "$lab-A1" "$BEFORE_BIN" "$BEFORE_ROOT" "$m" "$op" "$f" "$n"
  leg "$lab-B1" "$AFTER_BIN"  "$AFTER_ROOT"  "$m" "$op" "$f" "$n"
  leg "$lab-B2" "$AFTER_BIN"  "$AFTER_ROOT"  "$m" "$op" "$f" "$n"
  leg "$lab-A2" "$BEFORE_BIN" "$BEFORE_ROOT" "$m" "$op" "$f" "$n"
  leg "$lab-AA1" "$BEFORE_BIN" "$BEFORE_ROOT" "$m" "$op" "$f" "$n"
  leg "$lab-AA2" "$BEFORE_BIN" "$BEFORE_ROOT" "$m" "$op" "$f" "$n"
}
case_run open-fs   file-source open      test-data/poi/test-data/spreadsheet/54016.xls 60
case_run text-fs   file-source full-text test-data/poi/test-data/spreadsheet/54016.xls 40
case_run open-fac  facade-file open      test-data/poi/test-data/spreadsheet/54016.xls 60
case_run onecell   file-source one-cell  test-data/poi/test-data/spreadsheet/54016.xls 60
