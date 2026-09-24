#!/bin/bash
# usage: run-gates.sh gate...   (each gate into merge-scratch/gates/<gate>/)
cd /home/zhuhe/code/litchi-worktrees/merge-spec-gaps || exit 1
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/merge
export TMPDIR=/home/zhuhe/code/litchi-worktrees/merge-scratch/tmp
for g in "$@"; do
  out=/home/zhuhe/code/litchi-worktrees/merge-scratch/gates/$g
  rm -rf "$out"; mkdir -p "$out"
  LITCHI_GATE_OUTPUT=$out LITCHI_GATE_ONLY=$g python3 /home/zhuhe/code/litchi-worktrees/merge-scratch/run-integration.py > "$out/runner.out" 2>&1
  echo "$g: $(tail -1 $out/$g.log) $(tail -1 $out/runner.out)"
done
