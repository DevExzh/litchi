#!/usr/bin/env bash
# capture_opaque.sh <harness-binary> <output-dir> [harness-cwd]
#
# Before-capture for change 0568's OPAQUE-HEAVY GATE.
#
# The attribution binary (xls_source_attribution) takes only --input PATH; it has
# no corpus selector, so it cannot be pointed at the comments-opaque-heavy
# corpus. That corpus exists only inside litchi-perf-baseline, which builds it in
# memory for the xls_source_backed_* / xls_owned_source_* selectors. Those
# selectors DO publish per-sample deterministic logical counters under
# results[].source, including the exact quantity change 0568 moves:
# selected_worksheet_read_calls / selected_worksheet_read_bytes.
#
# So this script captures, per selector:
#   harness/<selector>.json                 --warmup 5 --samples 60, pinned, ASLR off
#   syscalls/<selector>.samples-{1,11}.*    strace -f -c pread64,statx isolation pair
#   <output-dir>/opaque-summary.json        extracted counters, written by the caller
set -euo pipefail
SCRATCH=/tmp/claude-1001/-home-zhuhe-code-litchi/4b20edf7-6b1c-40a9-9e79-3ca6a15219f0/scratchpad/measure-0568
if [ $# -lt 2 ]; then echo "usage: $0 <harness-binary> <output-dir> [harness-cwd]" >&2; exit 2; fi
HARNESS=$(readlink -f "$1"); mkdir -p "$2"; OUT=$(readlink -f "$2")
HARNESS_CWD=$(readlink -f "${3:-$SCRATCH/baseline-tree}")
CPU=${CPU:-17}; WARMUP=${WARMUP:-5}; SAMPLES=${SAMPLES:-60}
DEFAULT_SELECTORS="xls_source_backed_open xls_source_backed_open_list_worksheets xls_source_backed_open_one_cell xls_owned_source_open xls_owned_source_open_list_worksheets xls_owned_source_open_one_cell"
SELECTORS=${SELECTORS:-$DEFAULT_SELECTORS}
ARCH=$(uname -m); TMP="$SCRATCH/tmp"; mkdir -p "$TMP" "$OUT/harness" "$OUT/syscalls"
export TMPDIR="$TMP" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
[ -x "$HARNESS" ] || { echo "harness binary not executable: $HARNESS" >&2; exit 2; }

idle_percent() { mpstat -P "$CPU" 1 1 2>/dev/null | awk -v c="$CPU" '$1=="Average:" && $2==c {print $NF}'; }
IDLE_BEFORE=$(idle_percent); IDLE_BEFORE=${IDLE_BEFORE:-unknown}

run() { # receipt stdout argv...
  local rpath=$1 spath=$2; shift 2
  local started finished rc
  started=$(date -u +%Y-%m-%dT%H:%M:%S.%NZ)
  set +e; ( cd "$HARNESS_CWD" && "$@" ) > "$spath" 2> "${rpath%.receipt.json}.stderr"; rc=$?; set -e
  finished=$(date -u +%Y-%m-%dT%H:%M:%S.%NZ)
  python3 - "$rpath" "$rc" "$started" "$finished" "$HARNESS_CWD" "$@" <<'PY'
import json, sys
path, rc, started, finished, cwd, *argv = sys.argv[1:]
json.dump({"command": argv, "cwd": cwd, "returncode": int(rc), "started_utc": started,
           "finished_utc": finished,
           "env": {"TMPDIR": "scratchpad tmp", "RAYON_NUM_THREADS": "1", "OMP_NUM_THREADS": "1"}},
          open(path, "w"), indent=2)
PY
  if [ "$rc" -ne 0 ]; then echo "FAILED ($rc): $*" >&2; sed -n 1,20p "${rpath%.receipt.json}.stderr" >&2; exit 1; fi
}

for sel in $SELECTORS; do
  run "$OUT/harness/$sel.receipt.json" "$OUT/harness/$sel.stdout" \
    setarch "$ARCH" -R taskset -c "$CPU" "$HARNESS" \
      --warmup "$WARMUP" --samples "$SAMPLES" --case "$sel" --json "$OUT/harness/$sel.json"
  echo "harness $sel"
  for n in 1 11; do
    run "$OUT/syscalls/$sel.samples-$n.receipt.json" "$OUT/syscalls/$sel.samples-$n.stdout" \
      taskset -c "$CPU" strace -f -c -e trace=pread64,statx \
        -o "$OUT/syscalls/$sel.samples-$n.strace.txt" \
        "$HARNESS" --warmup 1 --samples "$n" --case "$sel" --json "$OUT/syscalls/$sel.samples-$n.json"
  done
  echo "syscall pair $sel"
done
IDLE_AFTER=$(idle_percent); IDLE_AFTER=${IDLE_AFTER:-unknown}
printf '{"cpu_idle_percent":{"before":"%s","after":"%s"},"pinned_cpu":%s,"warmup":%s,"samples":%s}\n' \
  "$IDLE_BEFORE" "$IDLE_AFTER" "$CPU" "$WARMUP" "$SAMPLES" > "$OUT/host-fragment.json"
echo "opaque capture done"
