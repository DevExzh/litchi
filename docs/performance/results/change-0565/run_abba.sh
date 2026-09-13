#!/usr/bin/env bash
# run_abba.sh <candidate-binary> <output-dir> [candidate-cwd]
#
# Change 0565 paired-latency matrix, reproducing the change-0560 protocol
# (docs/performance/results/change-0560/latency/abba-*.json) and the frozen
# 0565 plan: ten XLS selectors, one fresh `litchi-perf-baseline` child per
# leg and selector, `--warmup 5 --samples 60`, pinned to one CPU with
# `taskset -c 17`, ASLR disabled with `setarch $(uname -m) -R`, legs ordered
# A1 (baseline), B1 (candidate), B2 (candidate), A2 (baseline) per selector.
#
# Baseline legs run from the clean detached worktree so the harness records
# `git_revision` 6b13261e5 with `git_worktree_dirty: false`; candidate legs run
# from CANDIDATE_CWD (default: the main tree) because the harness reads
# `git rev-parse HEAD` / dirty state from its cwd. Pass a clean candidate
# worktree as the third argument if `tools/perf_abba_summary.py` (which
# refuses dirty reports) is wanted as a strict cross-check.
#
# Layout written under <output-dir>:
#   latency/abba-{A1,B1,B2,A2}-<selector>.json   harness reports (same layout as 0560)
#   latency/abba-<leg>-<selector>.receipt.json   command, cwd, env, exit code, timing
#   latency/abba-<leg>-<selector>.stderr         child stderr
#   matrix.json                                  binaries, hashes, CPU, ASLR, order, idle check
#   analysis.json                                produced by analyze.py at the end
#
# Environment overrides: CPU (17), WARMUP (5), SAMPLES (60), SELECTORS (space list),
# DRY_RUN=1 (allow the baseline binary to stand in as its own candidate, which turns
# the matrix into an A/A noise-floor run; the matrix records dry_run: true).
#
# NOTE on the candidate cwd: tools/perf_abba_summary.py refuses any leg whose
# environment.git_worktree_dirty is true, and requires every environment field
# except git_revision to be identical across the four legs - rustc_version
# included. So the candidate must be built with the same toolchain the baseline
# used (rustc 1.95.0, pinned by rust-toolchain.toml) and run from a CLEAN
# worktree, or the strict cross-check cannot be produced.
set -euo pipefail

SCRATCH=/tmp/claude-1001/-home-zhuhe-code-litchi/4b20edf7-6b1c-40a9-9e79-3ca6a15219f0/scratchpad/measure-0565
BASELINE_BIN="$SCRATCH/bin/baseline/litchi-perf-baseline"
BASELINE_CWD="$SCRATCH/baseline-tree"

if [ $# -lt 2 ]; then
  echo "usage: $0 <candidate-binary> <output-dir> [candidate-cwd]" >&2
  exit 2
fi
CANDIDATE_BIN=$(readlink -f "$1")
OUT=$(mkdir -p "$2" && readlink -f "$2")
CANDIDATE_CWD=${3:-/home/zhuhe/code/litchi}
CPU=${CPU:-17}
WARMUP=${WARMUP:-5}
SAMPLES=${SAMPLES:-60}
DEFAULT_SELECTORS="xls_source_backed_open xls_source_backed_open_list_worksheets xls_source_backed_open_one_cell xls_owned_source_open xls_owned_source_open_list_worksheets xls_owned_source_open_one_cell xls_semantic_one_cell xls_semantic_list_worksheets xls_semantic_one_edit_save xls_semantic_full_cell_scan"
SELECTORS=${SELECTORS:-$DEFAULT_SELECTORS}
ARCH=$(uname -m)
TMP="$SCRATCH/tmp"
mkdir -p "$OUT/latency" "$TMP"

[ -x "$BASELINE_BIN" ] || { echo "baseline binary missing: $BASELINE_BIN" >&2; exit 2; }
[ -x "$CANDIDATE_BIN" ] || { echo "candidate binary not executable: $CANDIDATE_BIN" >&2; exit 2; }
BASE_SHA=$(sha256sum "$BASELINE_BIN" | cut -d' ' -f1)
CAND_SHA=$(sha256sum "$CANDIDATE_BIN" | cut -d' ' -f1)
DRY_RUN=${DRY_RUN:-0}
if [ "$BASE_SHA" = "$CAND_SHA" ]; then
  if [ "$DRY_RUN" = "1" ]; then
    echo "DRY RUN: baseline and candidate binaries are identical ($BASE_SHA); this is an A/A noise-floor run, not evidence" >&2
  else
    echo "refusing to run: baseline and candidate binaries are identical ($BASE_SHA)" >&2
    exit 2
  fi
fi

# CPU occupancy check: refuse a busy pinned CPU (idle below 90%).
IDLE=$(mpstat -P "$CPU" 1 1 2>/dev/null | awk -v cpu="$CPU" '$1=="Average:" && $2==cpu {print $NF}')
IDLE=${IDLE:-unknown}
if [ "$IDLE" != "unknown" ] && awk -v i="$IDLE" 'BEGIN{exit !(i < 90.0)}'; then
  echo "CPU $CPU is busy (idle ${IDLE}%); choose another CPU with CPU=<n>" >&2
  exit 2
fi
ASLR_SYSCTL=$(cat /proc/sys/kernel/randomize_va_space)
# Prove setarch -R takes effect: personality bit 0x0040000 (ADDR_NO_RANDOMIZE).
PERSONALITY=$(setarch "$ARCH" -R bash -c 'cat /proc/self/personality')

START_UTC=$(date -u +%Y-%m-%dT%H:%M:%SZ)
ORDER=()
run_child() {
  local leg=$1 selector=$2 bin=$3 cwd=$4
  local name="abba-$leg-$selector"
  local report="$OUT/latency/$name.json"
  local started finished rc
  started=$(date -u +%Y-%m-%dT%H:%M:%S.%NZ)
  set +e
  ( cd "$cwd" && TMPDIR="$TMP" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1 \
      setarch "$ARCH" -R taskset -c "$CPU" "$bin" \
        --warmup "$WARMUP" --samples "$SAMPLES" --case "$selector" --json "$report" \
        > "$OUT/latency/$name.stdout" 2> "$OUT/latency/$name.stderr" )
  rc=$?
  set -e
  finished=$(date -u +%Y-%m-%dT%H:%M:%S.%NZ)
  python3 - "$OUT/latency/$name.receipt.json" "$name" "$leg" "$selector" "$bin" "$cwd" "$rc" "$started" "$finished" "$CPU" "$WARMUP" "$SAMPLES" "$ARCH" "$TMP" "$report" <<'PY'
import json, sys
(_, path, name, leg, selector, binary, cwd, rc, started, finished, cpu, warmup, samples, arch, tmp, report) = sys.argv
json.dump({
    "name": name, "leg": leg, "stage": "baseline" if leg.startswith("A") else "candidate",
    "selector": selector, "binary": binary, "cwd": cwd, "returncode": int(rc),
    "command": ["setarch", arch, "-R", "taskset", "-c", cpu, binary,
                "--warmup", warmup, "--samples", samples, "--case", selector, "--json", report],
    "env": {"TMPDIR": tmp, "RAYON_NUM_THREADS": "1", "OMP_NUM_THREADS": "1"},
    "started_utc": started, "finished_utc": finished,
}, open(path, "w"), indent=2)
PY
  if [ "$rc" -ne 0 ]; then
    echo "child $name failed (exit $rc); see $OUT/latency/$name.stderr" >&2
    exit 1
  fi
  ORDER+=("$name")
  echo "captured $name"
}

for selector in $SELECTORS; do
  run_child A1 "$selector" "$BASELINE_BIN" "$BASELINE_CWD"
  run_child B1 "$selector" "$CANDIDATE_BIN" "$CANDIDATE_CWD"
  run_child B2 "$selector" "$CANDIDATE_BIN" "$CANDIDATE_CWD"
  run_child A2 "$selector" "$BASELINE_BIN" "$BASELINE_CWD"
done
END_UTC=$(date -u +%Y-%m-%dT%H:%M:%SZ)

python3 - "$OUT/matrix.json" "$BASELINE_BIN" "$BASE_SHA" "$BASELINE_CWD" "$CANDIDATE_BIN" "$CAND_SHA" "$CANDIDATE_CWD" "$CPU" "$IDLE" "$ASLR_SYSCTL" "$PERSONALITY" "$WARMUP" "$SAMPLES" "$START_UTC" "$END_UTC" "$ARCH" "${ORDER[*]-}" "$SELECTORS" "$DRY_RUN" <<'PY'
import json, subprocess, sys
(_, path, bbin, bsha, bcwd, cbin, csha, ccwd, cpu, idle, aslr, personality, warmup, samples, start, end, arch, order, selectors, dry) = sys.argv
def rev(cwd):
    try:
        r = subprocess.run(["git", "-C", cwd, "rev-parse", "HEAD"], capture_output=True, text=True, check=True).stdout.strip()
        d = subprocess.run(["git", "-C", cwd, "status", "--porcelain"], capture_output=True, text=True, check=True).stdout.strip() != ""
        return {"revision": r, "dirty": d}
    except Exception as e:
        return {"revision": None, "dirty": None, "error": str(e)}
json.dump({
    "schema_version": 1, "record_kind": "litchi-perf-change-0565-abba-matrix",
    "protocol": {"legs": ["A1", "B1", "B2", "A2"], "interleaving": "per selector: A1, B1, B2, A2",
                 "warmup": int(warmup), "samples": int(samples), "pinned_cpu": int(cpu),
                 "aslr": f"setarch {arch} -R (personality {personality}; kernel randomize_va_space {aslr})",
                 "cpu_idle_percent_before_run": idle,
                 "child_env": {"TMPDIR": "scratchpad tmp (tmpfs, same filesystem as /tmp)", "RAYON_NUM_THREADS": "1", "OMP_NUM_THREADS": "1"},
                 "source": "docs/performance/results/change-0560 protocol; docs/performance/results/change-0565/plan.json"},
    "baseline": {"binary": bbin, "sha256": bsha, "cwd": bcwd, **rev(bcwd)},
    "candidate": {"binary": cbin, "sha256": csha, "cwd": ccwd, **rev(ccwd)},
    "selectors": selectors.split(), "order": order.split(),
    "dry_run": dry == "1",
    "dry_run_note": ("baseline and candidate are the SAME binary; every comparison in the analysis is an A/A "
                     "noise floor and is not evidence about any change") if dry == "1" else None,
    "started_utc": start, "finished_utc": end,
}, open(path, "w"), indent=2)
PY
echo "matrix written: $OUT/matrix.json"
python3 -B "$SCRATCH/analyze.py" --latency "$OUT/latency" --output "$OUT/analysis.json" --matrix "$OUT/matrix.json" || true
echo "analysis written: $OUT/analysis.json"
