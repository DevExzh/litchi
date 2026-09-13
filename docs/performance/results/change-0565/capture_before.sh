#!/usr/bin/env bash
# capture_before.sh <attribution-binary> <harness-binary> <output-dir> [harness-cwd]
#
# Deterministic-counter, syscall-isolation and sanity-latency capture for one
# stage of change 0565. The counters follow change 0560 and the syscall
# isolation follows change 0564 exactly.
#
# What it writes under <output-dir>:
#   attribution/<mode>-<op>.json   xls_source_attribution --warmups 20 --samples 100
#                                  (the change-0560 attribution-child setting), stdout JSON
#   syscalls/<stem>_<op>.samples-{1,11}.strace.txt
#                                  strace -f -c -e trace=pread64,statx at --samples 1 and 11,
#                                  one warmup each; (11 - 1) / 10 isolates one operation (0564)
#   traces/<stem>_<op>.pread.txt   strace -f -e trace=pread64 of one child at
#                                  --warmups 1 --samples 1, so two operations (0564)
#   harness/sanity-xls_source_backed_open.json
#                                  litchi-perf-baseline --warmup 5 --samples 60, pinned,
#                                  ASLR disabled. A p50 sanity reference, NOT the ABBA.
#   host.json                      host, CPU, ASLR, tool versions, binary and fixture identity
#   summary.json                   summarize_before.py, which reuses summarize()/counted()
#                                  from docs/performance/results/change-0564/summarize_open_reads.py
#
# Modes: file-source (the 0564 subject, a FileSource ReadAt wrapper over a staged
# file) and owned-readat (an in-memory owned source; it should show no pread64 at
# all, which is itself the evidence that the file-source penalty is filesystem I/O).
#
# Everything is pinned to CPU ${CPU:-17}; taskset wraps strace so the pin is
# inherited by the traced tree.
set -euo pipefail
SCRATCH=/tmp/claude-1001/-home-zhuhe-code-litchi/4b20edf7-6b1c-40a9-9e79-3ca6a15219f0/scratchpad/measure-0565
if [ $# -lt 3 ]; then
  echo "usage: $0 <attribution-binary> <harness-binary> <output-dir> [harness-cwd]" >&2
  exit 2
fi
ATTR=$(readlink -f "$1"); HARNESS=$(readlink -f "$2")
mkdir -p "$3"; OUT=$(readlink -f "$3")
HARNESS_CWD=$(readlink -f "${4:-$SCRATCH/baseline-tree}")
CPU=${CPU:-17}
INPUT=${INPUT:-/home/zhuhe/code/litchi/test-data/ole/xls/ConditionalFormattingSamples.xls}
ATTR_WARMUPS=${ATTR_WARMUPS:-20}
ATTR_SAMPLES=${ATTR_SAMPLES:-100}
HARNESS_WARMUP=${HARNESS_WARMUP:-5}
HARNESS_SAMPLES=${HARNESS_SAMPLES:-60}
MODES=${MODES:-"file-source owned-readat"}
OPERATIONS=${OPERATIONS:-"open list one-cell"}
ARCH=$(uname -m)
TMP="$SCRATCH/tmp"; mkdir -p "$TMP" "$OUT/attribution" "$OUT/syscalls" "$OUT/traces" "$OUT/harness"
export TMPDIR="$TMP" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1

[ -x "$ATTR" ] || { echo "attribution binary not executable: $ATTR" >&2; exit 2; }
[ -x "$HARNESS" ] || { echo "harness binary not executable: $HARNESS" >&2; exit 2; }
[ -r "$INPUT" ] || { echo "input not readable: $INPUT" >&2; exit 2; }

idle_percent() { mpstat -P "$CPU" 1 1 2>/dev/null | awk -v c="$CPU" '$1=="Average:" && $2==c {print $NF}'; }
IDLE_BEFORE=$(idle_percent); IDLE_BEFORE=${IDLE_BEFORE:-unknown}

write_receipt() { # path rc started finished argv...
  python3 - "$@" <<'PY'
import json, sys
path, rc, started, finished, *argv = sys.argv[1:]
json.dump({"command": argv, "returncode": int(rc), "started_utc": started, "finished_utc": finished,
           "env": {"TMPDIR": "scratchpad tmp", "RAYON_NUM_THREADS": "1", "OMP_NUM_THREADS": "1"}},
          open(path, "w"), indent=2)
PY
}
run() { # receipt-path stdout-path argv...
  local rpath=$1 spath=$2; shift 2
  local started finished rc
  started=$(date -u +%Y-%m-%dT%H:%M:%S.%NZ)
  set +e; "$@" > "$spath" 2> "${rpath%.receipt.json}.stderr"; rc=$?; set -e
  finished=$(date -u +%Y-%m-%dT%H:%M:%S.%NZ)
  write_receipt "$rpath" "$rc" "$started" "$finished" "$@"
  if [ "$rc" -ne 0 ]; then echo "FAILED ($rc): $*" >&2; sed -n 1,20p "${rpath%.receipt.json}.stderr" >&2; exit 1; fi
}

for mode in $MODES; do
  case $mode in
    file-source) stem=xls_file_source;;
    owned-readat) stem=xls_owned_readat;;
    *) stem=$(echo "$mode" | tr -c 'a-zA-Z0-9' '_');;
  esac
  for op in $OPERATIONS; do
    # 1. deterministic counters: the change-0560 attribution-child setting.
    run "$OUT/attribution/$mode-$op.receipt.json" "$OUT/attribution/$mode-$op.json" \
      taskset -c "$CPU" "$ATTR" --input "$INPUT" --mode "$mode" --operation "$op" \
        --warmups "$ATTR_WARMUPS" --samples "$ATTR_SAMPLES"
    echo "attribution $mode $op"
    # 2. the change-0564 strace -f -c isolation pair. --warmups 0 is rejected, so
    #    both children take one warmup and differ only in the sample count.
    for n in 1 11; do
      run "$OUT/syscalls/${stem}_${op}.samples-$n.receipt.json" "$OUT/syscalls/${stem}_${op}.samples-$n.json" \
        taskset -c "$CPU" strace -f -c -e trace=pread64,statx \
          -o "$OUT/syscalls/${stem}_${op}.samples-$n.strace.txt" \
          "$ATTR" --input "$INPUT" --mode "$mode" --operation "$op" --warmups 1 --samples "$n"
    done
    echo "syscall pair $mode $op"
    # 3. the change-0564 full pread64 capture: one warmup plus one sample, two operations.
    run "$OUT/traces/${stem}_${op}.pread.receipt.json" "$OUT/traces/${stem}_${op}.pread.json" \
      taskset -c "$CPU" strace -f -e trace=pread64 -o "$OUT/traces/${stem}_${op}.pread.txt" \
        "$ATTR" --input "$INPUT" --mode "$mode" --operation "$op" --warmups 1 --samples 1
    echo "pread trace $mode $op"
  done
done

# 4. harness sanity reference (NOT the ABBA): one pinned, ASLR-disabled child.
( cd "$HARNESS_CWD" && run "$OUT/harness/sanity-xls_source_backed_open.receipt.json" \
    "$OUT/harness/sanity-xls_source_backed_open.stdout" \
    setarch "$ARCH" -R taskset -c "$CPU" "$HARNESS" \
      --warmup "$HARNESS_WARMUP" --samples "$HARNESS_SAMPLES" --case xls_source_backed_open \
      --json "$OUT/harness/sanity-xls_source_backed_open.json" )
echo "harness sanity xls_source_backed_open"

IDLE_AFTER=$(idle_percent); IDLE_AFTER=${IDLE_AFTER:-unknown}

python3 - "$OUT/host.json" "$ATTR" "$HARNESS" "$INPUT" "$CPU" "$HARNESS_CWD" "$IDLE_BEFORE" "$IDLE_AFTER" "$ARCH" <<'PY'
import datetime, hashlib, json, os, platform, subprocess, sys
path, attr, harness, inp, cpu, cwd, idle_before, idle_after, arch = sys.argv[1:]
def sha(p): return hashlib.sha256(open(p, "rb").read()).hexdigest()
def out(*a):
    try: return subprocess.run(a, capture_output=True, text=True).stdout.strip()
    except Exception as exc: return f"<{exc}>"
strace_version = out("strace", "-V").splitlines()
json.dump({
  "schema_version": 1,
  "record_kind": "litchi-perf-change-0565-capture-host",
  "captured_utc": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
  "platform": platform.platform(), "kernel": out("uname", "-r"),
  "cpu_model": next((l.split(":", 1)[1].strip() for l in open("/proc/cpuinfo") if l.startswith("model name")), None),
  "logical_cpus": os.cpu_count(),
  "pinned_cpu": int(cpu),
  "cpu_idle_percent": {"before_capture": idle_before, "after_capture": idle_after,
                       "method": "mpstat -P <cpu> 1 1, the Average row's %idle column"},
  "loadavg": open("/proc/loadavg").read().strip(),
  "aslr": {"kernel_randomize_va_space": open("/proc/sys/kernel/randomize_va_space").read().strip(),
           "harness_child": f"setarch {arch} -R",
           "personality_under_setarch_R": out("setarch", arch, "-R", "bash", "-c", "cat /proc/self/personality"),
           "attribution_children": "not wrapped in setarch; they are counted, not timed"},
  "perf_event_paranoid": open("/proc/sys/kernel/perf_event_paranoid").read().strip(),
  "tools": {"strace": strace_version[0] if strace_version else None, "python": sys.version.split()[0],
            "mpstat": out("mpstat", "-V").splitlines()[0] if out("mpstat", "-V") else None},
  "tmpdir": {"path": os.environ.get("TMPDIR"), "filesystem": out("bash", "-c", "df -T \"$TMPDIR\" | tail -1 | awk '{print $2}'")},
  "attribution_binary": {"path": attr, "sha256": sha(attr)},
  "harness_binary": {"path": harness, "sha256": sha(harness), "cwd": cwd,
                     "git_revision": out("git", "-C", cwd, "rev-parse", "HEAD"),
                     "git_worktree_dirty": out("git", "-C", cwd, "status", "--porcelain") != ""},
  "input": {"path": inp, "bytes": os.stat(inp).st_size, "sha256": sha(inp)},
  "quiescence": "not established; the host runs unrelated concurrent workloads (another agent builds and tests in the main tree). "
                "The pinned CPU's idle percentage is recorded before and after the capture instead.",
}, open(path, "w"), indent=2)
PY
python3 -B "$SCRATCH/summarize_before.py" --capture "$OUT" --output "$OUT/summary.json"
