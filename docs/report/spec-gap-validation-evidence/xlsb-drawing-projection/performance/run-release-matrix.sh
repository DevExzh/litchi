#!/usr/bin/env bash
set -euo pipefail

# Run the complete supported-case matrix against one already-built binary.
# The Rust CLI intentionally accepts one case per process; these 26 cases are
# repeated in three fresh processes (78 invocations) and each output is a raw
# JSON report. The four backend lanes have deliberately different scopes.

repo_root=$(cd "$(dirname "${BASH_SOURCE[0]}")/../../../../.." && pwd)
target_dir=${CARGO_TARGET_DIR:?set CARGO_TARGET_DIR to the frozen release target}
binary=${XLSB_CRUD_BINARY:-"$target_dir/release/xlsb_crud"}
fixture=${XLSB_CRUD_FIXTURE:-"$repo_root/test-data/poi/test-data/spreadsheet/testVarious.xlsb"}
run_root=${XLSB_CRUD_RUN_ROOT:?set XLSB_CRUD_RUN_ROOT to a new output directory}
warmup=${XLSB_CRUD_WARMUP:-3}
samples=${XLSB_CRUD_SAMPLES:-30}

[[ -x "$binary" ]] || { echo "missing executable binary: $binary" >&2; exit 1; }
[[ -f "$fixture" ]] || { echo "missing fixture: $fixture" >&2; exit 1; }
[[ ! -e "$run_root" ]] || { echo "refusing to overwrite existing run directory: $run_root" >&2; exit 1; }
mkdir -p "$run_root"

cases=(
  open_identify
  worksheet_catalog
  selected_worksheet_cell
  full_stored_cell_scan
  noop_transaction_commit_save
  edit_one_existing_scalar_save
  edit_ceil_one_percent_existing_cells_save
)
owned_cases=(
  "${cases[@]}"
  full_text
)
source_cases=(
  open_identify
  worksheet_catalog
  selected_worksheet_cell
  full_stored_cell_scan
)

printf 'schema=\"litchi-xlsb-drawing-projection-run-v1\"\n' > "$run_root/commands.txt"
printf 'binary=%q\nfixture=%q\nwarmup=%q\nsamples=%q\nprocesses=3\n' \
  "$binary" "$fixture" "$warmup" "$samples" >> "$run_root/commands.txt"

run_case() {
  local backend=$1
  local process_id=$2
  local case_name=$3
  local output="$run_root/${backend}-p${process_id}-${case_name}.json"
  local log="$run_root/${backend}-p${process_id}-${case_name}.log"
  printf '%q --backend %q --case %q --fixture %q --warmup %q --samples %q --json %q\n' \
    "$binary" "$backend" "$case_name" "$fixture" "$warmup" "$samples" "$output" \
    >> "$run_root/commands.txt"
  "$binary" \
    --backend "$backend" \
    --case "$case_name" \
    --fixture "$fixture" \
    --warmup "$warmup" \
    --samples "$samples" \
    --json "$output" \
    > "$log" 2>&1
}

for process_id in 1 2 3; do
  for case_name in "${owned_cases[@]}"; do
    run_case owned "$process_id" "$case_name"
  done
  for case_name in "${cases[@]}"; do
    run_case owned_direct "$process_id" "$case_name"
  done
  for case_name in "${cases[@]}"; do
    run_case owned_without_drawings "$process_id" "$case_name"
  done
  for case_name in "${source_cases[@]}"; do
    run_case source_backed "$process_id" "$case_name"
  done
done

printf 'supported_backend_cases=26\nprocesses_per_case=3\ncompleted_invocations=78\n' > "$run_root/completed.txt"
