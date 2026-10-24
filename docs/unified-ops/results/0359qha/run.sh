#!/usr/bin/env bash
set -u -o pipefail

# Final-472b DOC body-toggle qualification packet.  This runner deliberately
# does not build: it reuses the source-bound binaries recorded by 0359qg.
# Native acceptance is 31 measured samples for each selected route and two
# complete baseline-candidate-candidate-baseline blocks for each route.  The
# DOC routes pass the batch-rounds setting directly to `time_rounds`, so the
# environment below is part of the sample-count contract.

ROOT=${ROOT:-/home/zhuhe/code/litchi-worktrees}
BASE_RUN=${BASE_RUN:-$ROOT/.coord/doc-toggle-472b-build-20261024}
SOURCE_SHORT=${SOURCE_SHORT:-472b}
RUN=${RUN:-$ROOT/.coord/doc-toggle-472b-performance-$SOURCE_SHORT}
PHASE=${PHASE:-native}
CPU=${CPU:?set CPU to the quiet pinned CPU}
ROUNDS=31
BATCH_ROUNDS=31
LEAF_INVOCATIONS=16
BATCH_INVOCATIONS=1
MIN_AGGREGATE_US=250000
PROFILE_ROUNDS=1
PROFILE_BATCH_ROUNDS=1
NATIVE_TIMEOUT_SECONDS=900
PROFILE_TIMEOUT_SECONDS=900
MAX_LOAD=1.0
MEMORY_MAX=20G
MEMORY_SWAP_MAX=0

CAND_SRC=$ROOT/.coord/final-api-capture-472b670547/source
BASE_SRC=$ROOT/.coord/final-api-capture-472b670547/doc-body-baseline
CAND_BIN=$BASE_RUN/target-candidate/release/examples/legacy_w5_source_probe
BASE_BIN=$BASE_RUN/target-baseline/release/examples/legacy_w5_source_probe
VALIDATOR=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/validate.py

EXPECTED_CAND_BODY=db01cfce38c13c6c806382ffede990fda2b3bbe4dd53f59c4c5cb9feba9777d3
EXPECTED_BASE_BODY=6dc26da8130659c212dd29f26499cdaba7c4849268583b44dfcd9413a36887a3
EXPECTED_CAND_REV=029721d34e7ad658e0064afd258cdb2b5a40e330110255b157aafb83f0d7b12a
EXPECTED_BASE_REV=b7155dc7c9aae6e318845b5073da4c0f275889125d38eae61df091aa2a2166df
EXPECTED_PROBE=b808e4dcba33dcafd4dff7dd22dc250e4170f2613839044fa8097263f5020965
EXPECTED_CAND_BIN=ba7042fd413c26764986b0ed9dcd79beca85e91231f0beb2f850550152e14877
EXPECTED_BASE_BIN=2d62dacae12b49214a086c109b506ca8eb69a33cf5e770627794efecdbcba810

FAIL=0
QUIET_FAILURE=0

die() { echo "0359qha: $*" >&2; exit 2; }
sha() { sha256sum "$1" | awk '{print $1}'; }
route_name() { echo "$1" | tr / _; }

expected_output() {
  case "$1:$2" in
    candidate:doc/leaf_write) echo 6bb97058e47f747c29e568395eea01db1226aa5ab8962087e1d78a1b9b82ff11 ;;
    baseline:doc/leaf_write) echo 6bb97058e47f747c29e568395eea01db1226aa5ab8962087e1d78a1b9b82ff11 ;;
    candidate:doc/batch1k_leaf_write) echo bd016a87f5d0015d6f8b910862eeeef9c0ce490033b4951c32ee5f403b7eebdf ;;
    baseline:doc/batch1k_leaf_write) echo 4a5bcc2888ab819488fbb5a1ced1b088010440bd8293f10556e0903993e448d3 ;;
    *) return 1 ;;
  esac
}

verify_source() {
  test "$(git -C "$CAND_SRC" rev-parse HEAD)" = 472b6705475b073e7fe08f6e11ee17fb1e29f838 || return 1
  test "$(git -C "$BASE_SRC" rev-parse HEAD)" = 472b6705475b073e7fe08f6e11ee17fb1e29f838 || return 1
  test -z "$(git -C "$CAND_SRC" status --porcelain=v1)" || return 1
  test "$(git -C "$BASE_SRC" status --porcelain=v1)" = $' M crates/litchi-doc/src/body_text.rs\n M crates/litchi-doc/src/tracked_revision/package.rs' || return 1
  test "$(sha "$CAND_SRC/crates/litchi-doc/src/body_text.rs")" = "$EXPECTED_CAND_BODY" || return 1
  test "$(sha "$BASE_SRC/crates/litchi-doc/src/body_text.rs")" = "$EXPECTED_BASE_BODY" || return 1
  test "$(sha "$CAND_SRC/crates/litchi-doc/src/tracked_revision/package.rs")" = "$EXPECTED_CAND_REV" || return 1
  test "$(sha "$BASE_SRC/crates/litchi-doc/src/tracked_revision/package.rs")" = "$EXPECTED_BASE_REV" || return 1
  test "$(sha "$CAND_SRC/crates/litchi-ops/examples/legacy_w5_source_probe.rs")" = "$EXPECTED_PROBE" || return 1
  test "$(sha "$BASE_SRC/crates/litchi-ops/examples/legacy_w5_source_probe.rs")" = "$EXPECTED_PROBE" || return 1
  test "$(sha "$CAND_SRC/Cargo.lock")" = 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1 || return 1
  test "$(sha "$BASE_SRC/Cargo.lock")" = 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1 || return 1
  test "$(sha "$CAND_BIN")" = "$EXPECTED_CAND_BIN" || return 1
  test "$(sha "$BASE_BIN")" = "$EXPECTED_BASE_BIN" || return 1
}

validate_json() {
  leg=$1; route=$2; samples=$3; json=$4
  digest=$(expected_output "$leg" "$route") || return 1
  python3 "$VALIDATOR" --json "$json" --scenario "$route" --samples "$samples" \
    --output-sha "$digest"
}

route_invocations() {
  case "$1" in
    doc/leaf_write) echo "$LEAF_INVOCATIONS" ;;
    doc/batch1k_leaf_write) echo "$BATCH_INVOCATIONS" ;;
    *) return 1 ;;
  esac
}

validate_aggregate() {
  leg=$1; route=$2; samples=$3; invocations=$4; dir=$5
  digest=$(expected_output "$leg" "$route") || return 1
  python3 "$VALIDATOR" --directory "$dir" --scenario "$route" \
    --samples "$samples" --invocations "$invocations" --output-sha "$digest" \
    --min-aggregate-us "$MIN_AGGREGATE_US"
}

require_profile_artifacts() {
  tool=$1; dir=$2
  case "$tool" in
    callgrind)
      test -s "$dir/result.callgrind" || return 1
      test -s "$dir/annotate.txt" || return 1
      ;;
    massif)
      test -s "$dir/result.massif" || return 1
      test -s "$dir/ms_print.txt" || return 1
      ;;
    *) return 1 ;;
  esac
}

snapshot() {
  file=$1
  {
    printf 'timestamp='; date -u +%Y-%m-%dT%H:%M:%SZ
    printf 'load='; awk '{print $1}' /proc/loadavg
    if ps -eo comm= | rg -q '^(cargo|rustc|ninja|make)$'; then
      echo 'busy=1'
    else
      echo 'busy=0'
    fi
  } >"$file"
}

watch_load() {
  marker=$1; file=$2
  : >"$file"
  while test -e "$marker"; do
    snapshot_line=$(awk '{print $1}' /proc/loadavg)
    busy=0
    ps -eo comm= | rg -q '^(cargo|rustc|ninja|make)$' && busy=1
    printf 'timestamp=%s load=%s busy=%s\n' "$(date -u +%Y-%m-%dT%H:%M:%SZ)" "$snapshot_line" "$busy" >>"$file"
    sleep 1
  done
}

check_quiet_file() {
  file=$1
  awk -v max="$MAX_LOAD" '
    /load=/ { split($0, a, "load="); split(a[2], b, " "); if ((b[1] + 0) > max) bad=1 }
    /busy=1/ { bad=1 }
    END { exit bad ? 1 : 0 }
  ' "$file"
}

native_leg() {
  kind=$1; route=$2; block=$3; ordinal=$4; leg=$5; bin=$6
  name=$(route_name "$route")
  dir="$RUN/native/$kind/$name/$block/$ordinal-$leg"
  mkdir -p "$dir" "$RUN/native-tmp/$kind/$name/$block/$ordinal-$leg"
  time_log="$dir/time.log"; marker="$dir/running"
  before="$dir/before.txt"; during="$dir/during.txt"; after="$dir/after.txt"
  invocations=$(route_invocations "$route") || return 1
  test ! -e "$dir/result-1.json" || return 1
  snapshot "$before"
  : >"$marker"
  watch_load "$marker" "$during" & watcher=$!
  for invocation in $(seq 1 "$invocations"); do
    json="$dir/result-$invocation.json"
    tmpdir="$RUN/native-tmp/$kind/$name/$block/$ordinal-$leg/$invocation"
    mkdir -p "$tmpdir"
    timeout --kill-after=30s "${NATIVE_TIMEOUT_SECONDS}s" \
      systemd-run --user --scope --quiet --collect \
        -p OOMPolicy=continue -p MemoryMax="$MEMORY_MAX" -p MemorySwapMax="$MEMORY_SWAP_MAX" -- \
        /usr/bin/time -f "$leg $name block=$block run=$ordinal invocation=$invocation maxrss_kb=%M wall=%e" \
        taskset -c "$CPU" env LITCHI_PROBE_ONLY="$route" \
          LITCHI_PROBE_BATCH_ROUNDS="$BATCH_ROUNDS" \
          TMPDIR="$tmpdir" "$bin" "$ROUNDS" \
        >"$json" 2>>"$time_log" || FAIL=1
  done
  rm -f "$marker"
  wait "$watcher" 2>/dev/null || true
  snapshot "$after"
  validate_aggregate "$leg" "$route" "$BATCH_ROUNDS" "$invocations" "$dir" || FAIL=1
  check_quiet_file "$before" || QUIET_FAILURE=1
  check_quiet_file "$during" || QUIET_FAILURE=1
  check_quiet_file "$after" || QUIET_FAILURE=1
}

profile_leg() {
  leg=$1; route=$2; tool=$3; bin=$4
  name=$(route_name "$route")
  dir="$RUN/allocation/$tool/$leg/$name"
  mkdir -p "$dir" "$RUN/profile-tmp/$tool/$leg/$name"
  json="$dir/result.json"
  test ! -e "$json" || return 1
  if test "$tool" = callgrind; then
    timeout --kill-after=30s "${PROFILE_TIMEOUT_SECONDS}s" \
      systemd-run --user --scope --quiet --collect \
        -p OOMPolicy=continue -p MemoryMax="$MEMORY_MAX" -p MemorySwapMax="$MEMORY_SWAP_MAX" -- \
        taskset -c "$CPU" env LITCHI_PROBE_ONLY="$route" \
          LITCHI_PROBE_BATCH_ROUNDS="$PROFILE_BATCH_ROUNDS" \
          TMPDIR="$RUN/profile-tmp/$tool/$leg/$name" valgrind --tool=callgrind --quiet \
            --callgrind-out-file="$dir/result.callgrind" "$bin" "$PROFILE_ROUNDS" \
        >"$json" 2>"$dir/time.log" || FAIL=1
    if test -s "$dir/result.callgrind"; then
      callgrind_annotate --auto=no --inclusive=yes --tree=calling \
        --show-percs=no --threshold=99 "$dir/result.callgrind" >"$dir/annotate.txt" || FAIL=1
    else
      FAIL=1
    fi
  else
    timeout --kill-after=30s "${PROFILE_TIMEOUT_SECONDS}s" \
      systemd-run --user --scope --quiet --collect \
        -p OOMPolicy=continue -p MemoryMax="$MEMORY_MAX" -p MemorySwapMax="$MEMORY_SWAP_MAX" -- \
        taskset -c "$CPU" env LITCHI_PROBE_ONLY="$route" \
          LITCHI_PROBE_BATCH_ROUNDS="$PROFILE_BATCH_ROUNDS" \
          TMPDIR="$RUN/profile-tmp/$tool/$leg/$name" valgrind --tool=massif --quiet \
            --pages-as-heap=no --massif-out-file="$dir/result.massif" "$bin" "$PROFILE_ROUNDS" \
        >"$json" 2>"$dir/time.log" || FAIL=1
    if test -s "$dir/result.massif"; then
      ms_print "$dir/result.massif" >"$dir/ms_print.txt" || FAIL=1
    else
      FAIL=1
    fi
  fi
  require_profile_artifacts "$tool" "$dir" || FAIL=1
  validate_json "$leg" "$route" "$PROFILE_ROUNDS" "$json" || FAIL=1
}

verify_source
source_status=$?
if test "$source_status" -ne 0; then
  die 'source, lock, overlay-path, or accepted binary hash mismatch'
fi
if test "$PHASE" = check; then
  echo '0359qha: source and binary identity checks passed'
  exit 0
fi
if test "$PHASE" = check-profile; then
  profile_tool=${PROFILE_TOOL:-}
  profile_dir=${PROFILE_DIR:-}
  test -n "$profile_tool" || die 'PROFILE_TOOL is required for check-profile'
  test -n "$profile_dir" || die 'PROFILE_DIR is required for check-profile'
  require_profile_artifacts "$profile_tool" "$profile_dir" ||
    die 'profile output or annotation is missing/invalid'
  echo '0359qha: profile artifact checks passed'
  exit 0
fi
test ! -e "$RUN" || die "RUN already exists; choose a fresh external root: $RUN"
mkdir -p "$RUN"
{
  printf 'source_short=%s\nrounds=%s\nbatch_rounds=%s\n' "$SOURCE_SHORT" "$ROUNDS" "$BATCH_ROUNDS"
  printf 'leaf_invocations=%s\nbatch_invocations=%s\nmin_aggregate_us=%s\n' \
    "$LEAF_INVOCATIONS" "$BATCH_INVOCATIONS" "$MIN_AGGREGATE_US"
  printf 'cpu=%s\nmemory_max=%s\nmemory_swap_max=%s\n' "$CPU" "$MEMORY_MAX" "$MEMORY_SWAP_MAX"
  printf 'candidate_commit=%s\n' "$(git -C "$CAND_SRC" rev-parse HEAD)"
  printf 'baseline_commit=%s\n' "$(git -C "$BASE_SRC" rev-parse HEAD)"
  printf 'candidate_body=%s\n' "$(sha "$CAND_SRC/crates/litchi-doc/src/body_text.rs")"
  printf 'baseline_body=%s\n' "$(sha "$BASE_SRC/crates/litchi-doc/src/body_text.rs")"
  printf 'candidate_revision=%s\n' "$(sha "$CAND_SRC/crates/litchi-doc/src/tracked_revision/package.rs")"
  printf 'baseline_revision=%s\n' "$(sha "$BASE_SRC/crates/litchi-doc/src/tracked_revision/package.rs")"
  printf 'probe=%s\n' "$(sha "$CAND_SRC/crates/litchi-ops/examples/legacy_w5_source_probe.rs")"
  printf 'lock=%s\n' "$(sha "$CAND_SRC/Cargo.lock")"
  printf 'candidate_binary=%s\n' "$(sha "$CAND_BIN")"
  printf 'baseline_binary=%s\n' "$(sha "$BASE_BIN")"
} >"$RUN/protocol.txt"

if test "$PHASE" = allocation || test "$PHASE" = all; then
  profile_leg candidate doc/leaf_write callgrind "$CAND_BIN"
  profile_leg baseline doc/leaf_write callgrind "$BASE_BIN"
  profile_leg candidate doc/batch1k_leaf_write callgrind "$CAND_BIN"
  profile_leg baseline doc/batch1k_leaf_write callgrind "$BASE_BIN"
  profile_leg candidate doc/leaf_write massif "$CAND_BIN"
  profile_leg baseline doc/leaf_write massif "$BASE_BIN"
  profile_leg candidate doc/batch1k_leaf_write massif "$CAND_BIN"
  profile_leg baseline doc/batch1k_leaf_write massif "$BASE_BIN"
fi

if test "$PHASE" = native || test "$PHASE" = all; then
  # Calibration legs use the final 31-sample settings and are not acceptance
  # samples.  Process wall time is recorded for diagnosis only: it includes
  # fixture/setup work and cannot establish the operation timing threshold.
  native_leg calibration doc/leaf_write 0 1 baseline "$BASE_BIN"
  native_leg calibration doc/leaf_write 0 2 candidate "$CAND_BIN"
  native_leg calibration doc/batch1k_leaf_write 0 1 baseline "$BASE_BIN"
  native_leg calibration doc/batch1k_leaf_write 0 2 candidate "$CAND_BIN"
  for block in 1 2; do
    native_leg abba doc/leaf_write "$block" 1 baseline "$BASE_BIN"
    native_leg abba doc/leaf_write "$block" 2 candidate "$CAND_BIN"
    native_leg abba doc/leaf_write "$block" 3 candidate "$CAND_BIN"
    native_leg abba doc/leaf_write "$block" 4 baseline "$BASE_BIN"
    native_leg abba doc/batch1k_leaf_write "$block" 1 baseline "$BASE_BIN"
    native_leg abba doc/batch1k_leaf_write "$block" 2 candidate "$CAND_BIN"
    native_leg abba doc/batch1k_leaf_write "$block" 3 candidate "$CAND_BIN"
    native_leg abba doc/batch1k_leaf_write "$block" 4 baseline "$BASE_BIN"
  done
fi

find "$RUN" -type f ! -name sha256.manifest -print0 | sort -z | xargs -0 sha256sum >"$RUN/sha256.manifest"
if test "$QUIET_FAILURE" -ne 0; then
  echo '0359qha: rejected: before/during/after host quiet/load check failed' >&2
  exit 3
fi
if test "$FAIL" -ne 0; then
  echo '0359qha: rejected: command, timeout, minimum-leg, or output check failed' >&2
  exit 4
fi
echo "0359qha: $PHASE packet complete"
