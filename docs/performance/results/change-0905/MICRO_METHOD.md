# 0905 separate native audit guards

The standalone `attribute_summary_probe` runs 12 fixed inputs under source,
compact and authored slice policies (36 jobs): 0/1/2/8/32/33/128 attributes,
valid/invalid xml:space, duplicate attributes, malformed tail and malformed
second attribute. Hex input and exact Debug outcomes accompany every child.
Every sample times 1,000 complete audit calls with black-boxed input/results.
Input construction, output formatting and last-result checking are outside the
timer; allocations and drops occurring within the batch remain included.

Six paired blocks alternate baseline/candidate order. Odd blocks also reverse
the flattened case/policy order. Each fresh CPU-12 child uses five warmup
batches and thirty measured batches per job. Twelve children therefore retain
12,960 measured batch samples. This population is separate from the Office
benefit gate and does not estimate whole-workflow speedup.

For each job, calculate the median of six candidate/baseline child-p50 ratios.
Bootstrap those six paired ratios with 10,000 draws, seed 905019, in sorted job
order; report sorted endpoints at indices 249 and 9749. A material adverse
regression is a paired-p50 median above 1.05 with lower endpoint above 1.0.
Any such result prevents adoption of this candidate. Median p95/p99 ratios
above 1.05 and per-role p50 max/min above 1.20 are additional review flags.
The small per-child sample count limits tail inference. Preserve all rows and
flags; do not repeat or discard slow children. Final adoption also requires
both unchanged primary Office benefit gates, semantic parity and quality.

`micro-plan.json` will bind this method, helper, identical probe source and both
probe executables before capture. `micro.py` retains raw TSV, per-child
receipts and exact input/outcome identity. Offline reduction does not require
live source/build roots. Only the coordinator executes it.
