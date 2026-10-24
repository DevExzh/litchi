#!/usr/bin/env bash
set -euo pipefail

ROOT=$(cd "$(dirname "$0")" && pwd)
RUNNER=$ROOT/run.sh

bash -n "$RUNNER"
rg -q '^ROUNDS=31$' "$RUNNER"
rg -q '^BATCH_ROUNDS=7$' "$RUNNER"
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
rg -q 'validate_json "\$leg" "\$route" "\$ROUNDS"' "$RUNNER"
rg -q 'validate_json "\$leg" "\$route" "\$PROFILE_ROUNDS"' "$RUNNER"
rg -q 'check_min_wall "\$time_log"' "$RUNNER"
rg -q '^MIN_LEG_SECONDS=0.250$' "$RUNNER"
printf '%s\n' '0359qha protocol static checks passed'
