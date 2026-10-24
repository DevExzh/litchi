#!/usr/bin/env bash
set -euo pipefail

ROOT=$(cd "$(dirname "$0")" && pwd)
RUNNER=$ROOT/run.sh
VALIDATOR=$ROOT/validate.py

bash -n "$RUNNER"
rg -q '^ROUNDS=31$' "$RUNNER"
rg -q '^BATCH_ROUNDS=31$' "$RUNNER"
rg -q 'for block in 1 2' "$RUNNER"
rg -q 'native_leg abba doc/leaf_write "\$block" 1 baseline' "$RUNNER"
rg -q 'native_leg abba doc/leaf_write "\$block" 2 candidate' "$RUNNER"
rg -q 'native_leg abba doc/leaf_write "\$block" 3 candidate' "$RUNNER"
rg -q 'native_leg abba doc/leaf_write "\$block" 4 baseline' "$RUNNER"
rg -q 'native_leg abba doc/batch1k_leaf_write "\$block" 1 baseline' "$RUNNER"
rg -q 'native_leg abba doc/batch1k_leaf_write "\$block" 4 baseline' "$RUNNER"
rg -q 'MemoryMax="\$MEMORY_MAX"' "$RUNNER"
rg -q 'MemorySwapMax="\$MEMORY_SWAP_MAX"' "$RUNNER"
rg -q 'timeout --kill-after=30s' "$RUNNER"
rg -q 'valgrind --tool=callgrind' "$RUNNER"
rg -q 'valgrind --tool=massif' "$RUNNER"
rg -q 'taskset -c "\$CPU"' "$RUNNER"
rg -q 'TMPDIR="\$RUN/native-tmp' "$RUNNER"
rg -q 'snapshot "\$before"' "$RUNNER"
rg -q 'watch_load "\$marker" "\$during"' "$RUNNER"
rg -q 'snapshot "\$after"' "$RUNNER"
rg -q 'test ! -e "\$RUN"' "$RUNNER"
rg -q 'validate_json "\$leg" "\$route" "\$BATCH_ROUNDS"' "$RUNNER"
rg -q 'validate_json "\$leg" "\$route" "\$PROFILE_ROUNDS"' "$RUNNER"
rg -q 'minimum_us=250000' "$RUNNER"
rg -q 'source_status=\$\?' "$RUNNER"
rg -q 'VALIDATOR=' "$RUNNER"
rg -q 'require_profile_artifacts' "$RUNNER"
rg -q 'PHASE.*check-profile' "$RUNNER"

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

# A bad early source hash must fail before any phase setup. This guards
# against Bash errexit being disabled by a function used in an || list.
BAD_RUNNER=$TMP/run-bad-hash.sh
sed 's/^EXPECTED_CAND_BODY=.*/EXPECTED_CAND_BODY=0000000000000000000000000000000000000000000000000000000000000000/' \
  "$RUNNER" >"$BAD_RUNNER"
chmod +x "$BAD_RUNNER"
if CPU=0 PHASE=check "$BAD_RUNNER" >/dev/null 2>&1; then
  echo 'bad early source hash was accepted' >&2
  exit 1
fi

# Profile receipts are executable requirements: a missing profiler output or
# a failed annotation must reject the packet, while complete artifacts pass.
PROFILE=$TMP/profile
mkdir -p "$PROFILE"
expect_profile_failure() {
  if CPU=0 PHASE=check-profile PROFILE_TOOL="$1" PROFILE_DIR="$PROFILE" \
    "$RUNNER" >/dev/null 2>&1; then
    echo "incomplete $1 profile was accepted" >&2
    exit 1
  fi
}
expect_profile_failure callgrind
printf '%s\n' counts >"$PROFILE/result.callgrind"
expect_profile_failure callgrind
printf '%s\n' annotation >"$PROFILE/annotate.txt"
if ! CPU=0 PHASE=check-profile PROFILE_TOOL=callgrind PROFILE_DIR="$PROFILE" \
  "$RUNNER" >/dev/null 2>&1; then
  echo 'complete callgrind profile was rejected' >&2
  exit 1
fi
rm -f "$PROFILE/result.callgrind" "$PROFILE/annotate.txt"
expect_profile_failure massif
printf '%s\n' snapshot >"$PROFILE/result.massif"
expect_profile_failure massif
printf '%s\n' summary >"$PROFILE/ms_print.txt"
if ! CPU=0 PHASE=check-profile PROFILE_TOOL=massif PROFILE_DIR="$PROFILE" \
  "$RUNNER" >/dev/null 2>&1; then
  echo 'complete massif profile was rejected' >&2
  exit 1
fi

EXPECTED_BATCH=bd016a87f5d0015d6f8b910862eeeef9c0ce490033b4951c32ee5f403b7eebdf
GOOD=$TMP/good.json
printf '[{"scenario":"doc/batch1k_leaf_write","samples":31,"median_us":250000,"min_us":250000,"output_sha256":"%s"}]\n' "$EXPECTED_BATCH" >"$GOOD"
python3 "$VALIDATOR" --json "$GOOD" --scenario doc/batch1k_leaf_write \
  --samples 31 --output-sha "$EXPECTED_BATCH" --min-operation-us 250000

# Missing receipts, short operation timing, and wrong sample counts all reject.
if python3 "$VALIDATOR" --json "$TMP/missing.json" --scenario doc/batch1k_leaf_write \
  --samples 31 --output-sha "$EXPECTED_BATCH" --min-operation-us 250000 >/dev/null 2>&1; then
  echo 'missing profile was accepted' >&2
  exit 1
fi
SHORT=$TMP/short.json
printf '[{"scenario":"doc/batch1k_leaf_write","samples":31,"median_us":249999,"min_us":249999,"output_sha256":"%s"}]\n' "$EXPECTED_BATCH" >"$SHORT"
if python3 "$VALIDATOR" --json "$SHORT" --scenario doc/batch1k_leaf_write \
  --samples 31 --output-sha "$EXPECTED_BATCH" --min-operation-us 250000 >/dev/null 2>&1; then
  echo 'subthreshold operation timing was accepted' >&2
  exit 1
fi
COUNT=$TMP/count.json
printf '[{"scenario":"doc/batch1k_leaf_write","samples":7,"median_us":250000,"min_us":250000,"output_sha256":"%s"}]\n' "$EXPECTED_BATCH" >"$COUNT"
if python3 "$VALIDATOR" --json "$COUNT" --scenario doc/batch1k_leaf_write \
  --samples 31 --output-sha "$EXPECTED_BATCH" --min-operation-us 250000 >/dev/null 2>&1; then
  echo 'short sample count was accepted' >&2
  exit 1
fi
printf '%s\n' '0359qha protocol static checks passed'
