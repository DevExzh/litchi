#!/usr/bin/env bash
# Final measurement sequence for change 0746 (all on CPU 20, nothing building).
S=/home/zhuhe/code/litchi-worktrees/scratch/0746
F=$S/final
mkdir -p $F
cd $S
echo "start $(date -u +%FT%TZ) $(uptime)" > $F/run.log
python3 tools/abba_probe.py $F/latency-probe 20 >> $F/run.log 2>&1 || echo "probe abba failed" >> $F/run.log
python3 tools/isolate_instructions.py $F/instructions 20 >> $F/run.log 2>&1 || echo "isolation failed" >> $F/run.log
mkdir -p $F/alloc
for f in 54016 xls-large; do for op in open number-plan number-source-backed number-generic string-generic comments-open visibility-open reader-open; do
  [ $f = xls-large ] && [ $op = string-generic ] && continue
  for leg in before after; do for run in 1 2; do
    ./bin/probe-$leg-alloc --input fixtures/$f.xls --operation $op --warmups 1 --samples 1 > $F/alloc/$f-$op-$leg-$run.json 2>&1 || echo "alloc failed $f $op $leg" >> $F/run.log
  done; done
done; done
python3 tools/abba_harness.py $F/latency-harness 20 >> $F/run.log 2>&1 || echo "harness abba failed" >> $F/run.log
python3 tools/isolate_harness.py $F/instructions-harness 20 xls_comments_eager_edit_save,xls_comments_eager_batch_edit_save,xls_comments_source_backed_batch_edit_save,xls_visibility_eager_edit_save,xls_visibility_eager_batch_edit_save,xls_semantic_noop_edit_save,xls_semantic_one_edit_save >> $F/run.log 2>&1 || echo "harness isolation failed" >> $F/run.log
echo "end $(date -u +%FT%TZ) $(uptime)" >> $F/run.log
echo "final run done" >> $F/run.log
