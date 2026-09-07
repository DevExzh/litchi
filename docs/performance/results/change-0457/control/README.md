# Change 0457 control evidence

This directory contains the owned current-revision `odp_existing_append_lifecycle` control: six forward lanes in R1, six reverse lanes in R2, 30 retained samples, three warmups, CPU 2, and one worker. It is a descriptive control only; it does not authorize a before/after, causal, bounded-memory, physical-I/O, scaling, or speedup claim.

The first `R1/formal` lane is retained as a failed attempt. The workload and report were valid, but the original copied oracle required `rustflags: -Cforce-frame-pointers=yes` while the preserved baseline report records `rustflags: null`. `protocol-r1.json` and `oracle/verify-report-r1.py` amend only that expectation, document that no frame-pointer build claim is made, and retain the original semantic oracle protocol. The amended protocol covers the R1 retry and the R2 current-build control.

Run the authenticated pre-cleanup verification while the copied binaries still exist:

```bash
python3 -B docs/performance/results/change-0457/control/verify.py \
  --protocol docs/performance/results/change-0457/control/protocol-r1.json \
  --attempt formal-r1 --precleanup
```

After binary cleanup, replay the retained evidence without `--precleanup`:

```bash
python3 -B docs/performance/results/change-0457/control/verify.py \
  --protocol docs/performance/results/change-0457/control/protocol-r1.json \
  --attempt formal-r1
```

The verifier is portable when the complete `change-0457` directory is copied: it reads artifacts from the copied bundle, while authenticated build-receipt `cwd` plus `docs/performance/results/change-0457/control` reconstructs the original absolute workload and oracle paths recorded in each receipt.

```bash
cp -a docs/performance/results/change-0457 /tmp/litchi-goal-0457-replay
python3 -B /tmp/litchi-goal-0457-replay/control/verify.py \
  --protocol /tmp/litchi-goal-0457-replay/control/protocol-r1.json \
  --attempt formal-r1
```

Derive exact per-lane vectors and descriptive p50/p95/p99 summaries, including RSS, sink counters, and allocator heap-growth/peak metrics:

```bash
python3 -B docs/performance/results/change-0457/control/derive.py \
  --protocol docs/performance/results/change-0457/control/protocol-r1.json \
  --attempt formal-r1 --write \
  --output docs/performance/results/change-0457/control/summary.json
```

The extractor defines p50 as the average of the two middle sorted values and p95/p99 as nearest-rank `ceil(p*n)` values. These are recomputed from retained vectors and are not silently mixed with report-embedded percentile fields.
