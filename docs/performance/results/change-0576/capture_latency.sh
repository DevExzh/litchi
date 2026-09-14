#!/usr/bin/env bash
# capture_latency.sh <before-binary> <after-binary> <out-dir>
#
# A/B/B/A wall-clock capture for change 0576. Each leg is a separate child
# process; the two directions bracket each other so that any monotonic drift in
# the host shows up as a disagreement between them.
#
# The per-leg driver is change 0574's capture_counters.sh, unchanged, so the
# children are launched exactly as that record's attribution children were:
# `setarch x86_64 -R` (ASLR off), `taskset -c $CPU`, single-threaded, TMPDIR in
# the session scratchpad.
set -euo pipefail
BEFORE=$(readlink -f "$1")
AFTER=$(readlink -f "$2")
mkdir -p "$3"; OUT=$(readlink -f "$3")
HERE=$(cd "$(dirname "$0")" && pwd)
DRIVER="$HERE/../change-0574/capture_counters.sh"
REPO=${REPO:-/home/zhuhe/code/litchi}

export CPU=${CPU:-17}
export MODES=${MODES:-"owned-readat file-source"}
export W=${W:-20}
export S=${S:-100}

fixtures=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls"
)

# The driver is invoked once per mode so that one unavailable cell cannot lose
# the rest of that fixture's matrix. `54016.xls` has fewer worksheets than the
# driver's fixed `--worksheet-index 1`, so its `one-cell` operation fails with
# `WorksheetNotFound("1")` and is simply absent from the results; `open` is the
# operation this change acts on and it is captured everywhere.
for round in a1 b1 b2 a2; do
  case "$round" in
    a1|a2) BIN="$BEFORE" ;;
    b1|b2) BIN="$AFTER" ;;
  esac
  for entry in "${fixtures[@]}"; do
    stem=${entry%%:*}
    path=${entry#*:}
    for mode in $MODES; do
      INPUT="$path" MODES="$mode" "$DRIVER" "$BIN" "$OUT/$round/$stem" >/dev/null 2>&1 || true
    done
    echo "captured $round/$stem"
  done
done
