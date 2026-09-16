#!/bin/bash
set -e
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0654
W=/home/zhuhe/code/litchi-worktrees/scratch-0654
CASE=xlsx_source_backed_cell_values_one_edit_save
cd /home/zhuhe/code/litchi
for leg in before after; do
  for n in 1 3; do
    taskset -c 9 valgrind --tool=callgrind \
      --callgrind-out-file="$W/cg/$leg-n$n.out" --cache-sim=no --branch-sim=no \
      "$W/bin/lpb-$leg" --case "$CASE" --warmup 0 --samples "$n" > "$W/cg/$leg-n$n.json" 2> "$W/cg/$leg-n$n.stderr"
    echo "== $leg n=$n =="
    python3 "$S/audit-call-counts.py" "$W/cg/$leg-n$n.out" \
      xml_minifier::audit::verify_authored \
      xml_minifier::audit::verify_source \
      xml_minifier::audit::verify_with_policy \
      xml_minifier::audit::package::is_xml_part
    grep -m1 '^summary:' "$W/cg/$leg-n$n.out"
    callgrind_annotate --inclusive=yes --threshold=100 "$W/cg/$leg-n$n.out" \
      | grep -E 'write_topology_to_stream \[|verify_authored \[|verify_source \[|verify_with_policy \[|validate_overlay_xml \[|validate_source_part_xml \[' || true
  done
done
