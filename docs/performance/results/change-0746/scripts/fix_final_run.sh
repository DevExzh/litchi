#!/usr/bin/env bash
# Review-fix measurement sequence for change 0746 (CPU 20, nothing building).
S=/home/zhuhe/code/litchi-worktrees/scratch/0746
F=$S/fix-final
mkdir -p $F
cd $S
CASES="54016.xls:open:3:20:open_ns;54016.xls:comments-open:3:20:open_ns;54016.xls:visibility-open:3:20:open_ns;54016.xls:number-generic:3:20:commit_ns;54016.xls:number-source-backed:3:20:commit_ns;54016.xls:reader-open:3:20:open_ns;sparse-bands-minimal.xls:open:2:10:open_ns;sparse-bands-minimal.xls:comments-open:2:10:open_ns;sparse-bands-minimal.xls:visibility-open:2:10:open_ns;sparse-bands-minimal.xls:reader-open:2:10:open_ns;xls-large.xls:open:5:60:open_ns;xls-large.xls:number-generic:5:60:commit_ns;xls-large.xls:reader-open:5:60:open_ns"
echo "start $(date -u +%FT%TZ) $(uptime)" > $F/run.log
python3 tools/abba_legs.py $F/base-vs-final 20 $S/bin/probe-before $S/bin/probe-after "$CASES" >> $F/run.log 2>&1 || echo "abba base failed" >> $F/run.log
echo "mid $(date -u +%FT%TZ) $(uptime)" >> $F/run.log
python3 tools/abba_legs.py $F/banded-vs-final 20 $S/bin/probe-prefix $S/bin/probe-after "$CASES" >> $F/run.log 2>&1 || echo "abba banded failed" >> $F/run.log
echo "isolation $(date -u +%FT%TZ) $(uptime)" >> $F/run.log
python3 tools/isolate_legs.py $F/isolation.json 20 "base=$S/bin/probe-before,banded=$S/bin/probe-prefix,final=$S/bin/probe-after" "sparse-bands-minimal.xls:open:2:8;sparse-bands-minimal.xls:comments-open:2:8;sparse-bands-minimal.xls:visibility-open:2:8;sparse-bands-minimal.xls:reader-open:2:8;54016.xls:open:4:24;54016.xls:comments-open:4:24;54016.xls:visibility-open:4:24;54016.xls:number-generic:4:24;54016.xls:number-source-backed:4:24;54016.xls:reader-open:4:24;xls-large.xls:open:4:24;xls-large.xls:number-generic:4:24;xls-large.xls:reader-open:4:24" >> $F/run.log 2>&1 || echo "isolation failed" >> $F/run.log
mkdir -p $F/alloc
for f in sparse-bands-minimal 54016 xls-large; do for op in open comments-open visibility-open reader-open; do for leg in before prefix after; do for run in 1 2; do
  ./bin/probe-$leg-alloc --input fixtures/$f.xls --operation $op --warmups 1 --samples 1 > $F/alloc/$f-$op-$leg-$run.json 2>&1 || echo "alloc failed $f $op $leg" >> $F/run.log
done; done; done; done
echo "end $(date -u +%FT%TZ) $(uptime)" >> $F/run.log
echo "fix final run done" >> $F/run.log
# Commit-operation allocation counts, run afterwards (deterministic, so not timing-sensitive):
for f in 54016 xls-large; do ops="number-plan number-source-backed number-generic"; [ $f = 54016 ] && ops="$ops string-generic"; for op in $ops; do for leg in before prefix after; do for run in 1 2; do
  taskset -c 20 ./bin/probe-$leg-alloc --input fixtures/$f.xls --operation $op --warmups 1 --samples 1 > $F/alloc/$f-$op-$leg-$run.json 2>&1 || echo "alloc failed $f $op $leg"
done; done; done; done
