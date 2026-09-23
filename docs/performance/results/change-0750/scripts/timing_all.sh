#!/bin/bash
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
BEFORE=$S/bin/litchi-perf-baseline.before AFTER=$S/bin/litchi-perf-baseline.after OUT=$S/timing CORE=24 $S/scripts/run_abba.sh
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0750-before/release/litchi-audit-probe-0750 AFTER=/home/zhuhe/code/litchi-worktrees/targets/0750-census-after/release/litchi-audit-probe-0750 PARTS=$S/parts OUT=$S/probe-timing CORE=24 $S/scripts/run_probe_abba.sh
echo alldone >> $S/timing-status.txt
