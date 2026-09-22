# Prospective 0737 harness qualification

Production is unchanged from accepted 0734 and restored 0735. The 0736 retrospective
found repeatable sample-position bands and full oracle witnesses retained between
owner calls. This experiment measures sensitivity to harness lifecycle, not an
optimization of production or the cause of 0735 secondary regression.

The original 0735 probe is rebuilt independently as archive control. A revised
single binary offers legacy, strict-retained and strict-drained. Both strict
arms execute the identical full oracle and compact receipt construction, validate
warmups, drain full warmup witnesses and differ only in retaining the full measured
witness until the last owner timer. Compact receipts remain live in both arms.
The legacy-a versus legacy-b labels run identical code and arguments as a noise
control. Archive versus revised legacy detects combined build/codegen/observer
changes; legacy versus strict-retained changes warmup validation AND compact
receipt work, so is not pure warmup attribution.

Run both fixtures, CPU 12, serial pinned processes, exact fixed rotated schedule.
108 native processes include 90 with 50 samples / 3 warmups and 18 fresh-child controls
with 1 sample / 3 warmups. Allocation lane: 30 processes,1 sample, 0 or 3 warmups. Native and
allocation results remain separate. Full owner output and every visible oracle
witness must equal sealed 0735; all eight corruption controls must still reject.

Report p50, mean, p95, p99, max and all >5% review flags; bootstrap paired process
ratios with seed 7337 and 10,000 resamples. Three descriptive windows (first 10,
middle 30, last 10) and per-process last 10 / first 10 drift retain sample order, never
replace full-window gates. Fresh children yield 18 observations from 18 processes,
not independent samples pooled with 50-sample processes. Allocator counters are
boundary-relative owner measures, not total probe heap or RSS. Full witness counts
prove retention policy but do not measure retained bytes. No allocation cause or
cache cause inferred. No candidate reinstatement authorized by this matrix.
