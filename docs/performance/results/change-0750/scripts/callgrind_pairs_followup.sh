#!/bin/bash
# Change 0750 follow-up: callgrind isolation pairs (0654's method) for the two
# re-measured cases, base against the expanded-name fix. Each harness binary
# runs each case at --samples 1 and --samples 3 with no warmup; half the
# difference is one sample. scripts/cg_attribute.py reads the auditor's entry
# points from these profiles (the harness keeps their symbols).
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
mkdir -p $S/cg-followup && cd $S/cg-followup
for leg in before after; do
  for spec in "sb_one_dense|--case xlsx_source_backed_cell_values_one_edit_save --xlsx-cell-crud-shape dense-sparse" \
              "docx_sb_one|--case docx_source_backed_one_edit_save"; do
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
