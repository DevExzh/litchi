#!/usr/bin/env bash
# change 0635: published-package sha256 per (case, shape), before vs after.
set -uo pipefail
STAGE=/home/zhuhe/code/litchi-worktrees/targets/0635-staged
OUT="${1:?usage: output-hashes.sh OUTDIR}"
mkdir -p "$OUT"
CASES="xlsx_source_backed_cell_values_one_edit_save,xlsx_source_backed_cell_values_one_percent_edit_save,xlsx_source_backed_cell_values_batch_edit_save,xlsx_source_backed_cell_values_multi_sheet_edit_save,xlsx_source_backed_managed_cell_values_one_edit_save,xlsx_source_backed_managed_cell_values_one_percent_edit_save,xlsx_source_backed_cell_clear_edit_save,xlsx_source_backed_cell_remove_edit_save"
SHAPES="medium,dense-sparse,noncompact,vendor-extension"
PRODUCER="xlsx_producer_medium_source_one_edit_save,xlsx_producer_dense_source_one_edit_save"
for leg in before after; do
  bin="$STAGE/$leg-litchi-perf-baseline"
  dir=/home/zhuhe/code/litchi-worktrees/before-c7326f680
  [ "$leg" = after ] && dir=/home/zhuhe/code/litchi-worktrees/0635
  ( cd "$dir" && taskset -c 11 "$bin" --case "$CASES" --xlsx-cell-crud-shape "$SHAPES" \
      --warmup 0 --samples 1 --json "$OUT/$leg-cells.json" >/dev/null 2>&1 )
  ( cd "$dir" && taskset -c 11 "$bin" --case "$PRODUCER" \
      --warmup 0 --samples 1 --json "$OUT/$leg-producer.json" >/dev/null 2>&1 )
done
python3 - "$OUT" <<'PY'
import json, sys, os
out = sys.argv[1]
rows = {}
for leg in ['before', 'after']:
    for part in ['cells', 'producer']:
        path = os.path.join(out, f'{leg}-{part}.json')
        with open(path) as handle:
            data = json.load(handle)
        for record in data['results']:
            key = (record['case'], record['corpus'].get('shape') or '-')
            rows.setdefault(key, {})[leg] = record['output_sha256']
same = sum(1 for v in rows.values() if v.get('before') == v.get('after'))
lines = [f"# change 0635: published package sha256 per (case, shape), before vs after",
         f"# {len(rows)} pairs compared, {same} identical, {len(rows) - same} differing", ""]
for key in sorted(rows):
    for leg in ['before', 'after']:
        lines.append(f"{leg} {key[0]} {key[1]} {rows[key][leg]}")
with open(os.path.join(out, 'output-hashes.txt'), 'w') as handle:
    handle.write("\n".join(lines) + "\n")
print(lines[1])
PY
