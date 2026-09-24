#!/usr/bin/env bash
set -euo pipefail

# This script is prepared for the post-review capture window.  It is gated so
# merely checking the artifact cannot start timing samples.
if [[ "${RUN_CAPTURE_AFTER_REVIEW:-0}" != 1 ]]; then
  echo 'capture deferred: set RUN_CAPTURE_AFTER_REVIEW=1 after PPTX review clears'
  exit 0
fi

SOURCE=/var/tmp/litchi-performance-full-baseline-capture-control-20260910
ARTIFACT=/var/tmp/litchi-performance-full-baseline-20260910/capture-control
TARGET="$ARTIFACT/cargo-target"
LOCK_PATCH=/var/tmp/litchi-performance-full-baseline-20260909/Cargo.lock.generated.patch
CPU_LIST=2
SAMPLES=15
WARMUP=3
RUSTFLAGS_VALUE='-D warnings -D deprecated'
NORMAL_BIN="$TARGET/release/litchi-perf-baseline"
ALLOC_BIN="$TARGET/release/litchi-perf-baseline-alloc"
NORMAL_REPORT="$ARTIFACT/full-normal.json"
NORMAL_CATALOG="$ARTIFACT/full-normal.corpus-manifest-v2.json"
ALLOC_REPORT="$ARTIFACT/opc-file-allocator.json"
ALLOC_CATALOG="$ARTIFACT/opc-file-allocator.corpus-manifest-v2.json"

mkdir -p "$ARTIFACT"
LOCK_BACKUP="$ARTIFACT/control-Cargo.lock.clean"
cp "$SOURCE/tools/perf-baseline/Cargo.lock" "$LOCK_BACKUP"
restore_lock() {
  cp "$LOCK_BACKUP" "$SOURCE/tools/perf-baseline/Cargo.lock"
}
trap restore_lock EXIT HUP INT TERM

# Cargo.lock in the frozen source commit predates two manifest dependencies.
# Apply the reviewed minimal lock correction only for the build, then restore
# the clean source before the report binaries run so git_worktree_dirty=false.
git -C "$SOURCE" diff --quiet -- tools/perf-baseline/Cargo.lock
git -C "$SOURCE" apply "$LOCK_PATCH"
CARGO_ENV=(
  RUSTUP_HOME=/tmp/litchi-spec-gap-rustup
  CARGO_TARGET_DIR="$TARGET"
  CARGO_BUILD_JOBS=4
  CARGO_INCREMENTAL=0
  CARGO_PROFILE_RELEASE_DEBUG=0
  RUSTFLAGS="$RUSTFLAGS_VALUE"
  LC_ALL=C.UTF-8
)
env "${CARGO_ENV[@]}" cargo +1.95.0 build --release --locked \
  --manifest-path "$SOURCE/tools/perf-baseline/Cargo.toml" \
  --bin litchi-perf-baseline 2>&1 | tee "$ARTIFACT/build-normal.log"
env "${CARGO_ENV[@]}" cargo +1.95.0 build --release --locked \
  --features allocator-metrics \
  --manifest-path "$SOURCE/tools/perf-baseline/Cargo.toml" \
  --bin litchi-perf-baseline-alloc 2>&1 | tee "$ARTIFACT/build-allocator.log"
restore_lock
git -C "$SOURCE" diff --quiet -- tools/perf-baseline/Cargo.lock

record_pinned_snapshot() {
  local path="$1"
  taskset --cpu-list "$CPU_LIST" env SNAPSHOT_PATH="$path" bash -c '
    {
      echo "snapshot_kind=pre_run_contention"
      date --iso-8601=seconds
      taskset -pc $$ 2>&1
      echo "loadavg="; cat /proc/loadavg
      echo "cpus_allowed_list="; sed -n "s/^Cpus_allowed_list:[[:space:]]*//p" /proc/self/status
      echo "online_cpus="; cat /sys/devices/system/cpu/online
      echo "top_processes="
      ps -eo pid,ppid,psr,pcpu,pmem,stat,comm --sort=-pcpu | head -n 26
    } > "$SNAPSHOT_PATH"
  '
}

record_pinned_snapshot "$ARTIFACT/normal-pre-run-contention.txt"
taskset --cpu-list "$CPU_LIST" env \
  RUSTFLAGS="$RUSTFLAGS_VALUE" LC_ALL=C.UTF-8 \
  "$NORMAL_BIN" \
  --warmup "$WARMUP" --samples "$SAMPLES" \
  --filesystem-cache warm,cold-requested \
  --json "$NORMAL_REPORT" \
  --corpus-manifest "$NORMAL_CATALOG" \
  2>&1 | tee "$ARTIFACT/normal-run.log"

record_pinned_snapshot "$ARTIFACT/allocator-pre-run-contention.txt"
taskset --cpu-list "$CPU_LIST" env \
  RUSTFLAGS="$RUSTFLAGS_VALUE" LC_ALL=C.UTF-8 \
  "$ALLOC_BIN" \
  --warmup "$WARMUP" --samples "$SAMPLES" \
  --case opc_file_eager_open \
  --filesystem-cache warm,cold-requested \
  --json "$ALLOC_REPORT" \
  --corpus-manifest "$ALLOC_CATALOG" \
  2>&1 | tee "$ARTIFACT/allocator-run.log"

python3 - "$NORMAL_REPORT" "$ALLOC_REPORT" <<'PY'
import json
import sys
from pathlib import Path
normal = json.loads(Path(sys.argv[1]).read_text())
allocator = json.loads(Path(sys.argv[2]).read_text())
assert len(normal["results"]) == 201, len(normal["results"])
assert len(allocator["results"]) == 2, len(allocator["results"])
assert normal["configuration"]["samples_per_case"] == 15
assert normal["configuration"]["warmup_iterations_per_case"] == 3
assert allocator["configuration"]["samples_per_case"] == 15
assert allocator["configuration"]["warmup_iterations_per_case"] == 3
assert allocator["tool"]["binary"] == "litchi-perf-baseline-alloc"
assert allocator["tool"]["instrumentation"] == "system_allocator_operation_scoped"
print("normal_results=201 allocator_results=2 samples=15 warmup=3")
PY

sha256sum "$NORMAL_BIN" "$ALLOC_BIN" "$NORMAL_REPORT" "$NORMAL_CATALOG" "$ALLOC_REPORT" "$ALLOC_CATALOG" > "$ARTIFACT/output-sha256.txt"
test -z "$(git -C "$SOURCE" status --porcelain)"
printf '%s\n' 'capture complete; source worktree clean'
