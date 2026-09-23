#!/bin/bash
# Change 0750: callgrind isolation pairs (0654's method): each harness binary
# runs each case at --samples 1 and --samples 3 with no warmup; half the
# difference of the program totals is one sample (the timed operation plus the
# harness's untimed per-sample preparation and checks), which cancels corpus
# construction and the harness's own gates. The harness binaries have
# no symbols for the audit (LTO inlines it), so only program totals are used;
# the audit alone is counted by the audit probe.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
mkdir -p $S/cg-pairs && cd $S/cg-pairs
for leg in before after; do
  for spec in "sb_one_medium|--case xlsx_source_backed_cell_values_one_edit_save --xlsx-cell-crud-shape medium" \
              "sb_one_dense|--case xlsx_source_backed_cell_values_one_edit_save --xlsx-cell-crud-shape dense-sparse" \
              "docx_sb_one|--case docx_source_backed_one_edit_save" \
              "pptx_sb_one|--case pptx_source_backed_one_edit_save"; do
    label=${spec%%|*}; args=${spec#*|}
    for s in 1 3; do
      # shellcheck disable=SC2086
      taskset -c 24 valgrind --tool=callgrind --callgrind-out-file=cg-$leg-$label-s$s.out \
        $S/bin/litchi-perf-baseline.$leg $args --samples $s --warmup 0 --json run-$leg-$label-s$s.json \
        > log-$leg-$label-s$s.txt 2>&1
      echo "$leg $label s$s exit $?" >> status.txt
    done
  done
done
echo done >> status.txt
