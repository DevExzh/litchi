# Final byte-function performance capture

The source-frozen capture completed successfully against candidate freeze
[`performance/gates/freeze.json`](performance/gates/freeze.json). The baseline
is `3844f235bac545ff0ae1580b97612883c1fd9f89`; the candidate and baseline
used the frozen gate lock SHA256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`, and the
contract SHA256 is
`88dafc6f0d111672e724b4238289afc0a17d879b856059ac3886ca60bbc99db5`.

The run used three warmups and fifteen fresh child samples in both
`evaluate` and `parse-evaluate`. It retained 840 baseline records (28 matched
controls) and 2,520 candidate records (84 cases), including all seven byte
functions over tiny, long Unicode, long ASCII, borrowed 64-cell references,
matrix broadcast, typed refusal, cancellation, resource refusal, worst-case
search, REPLACEB growth, and the three projected statistical lanes. The
independent UTF-8 oracle preflight passed every case with exact output,
shape, and reference-read checks. `verify.py` passed with the frozen source
manifest and unchanged profile-input hashes.

Cancellation lanes use four internal repeats. Their raw receipts record one
successful resolver read per child before sticky cancellation. The displayed
normalized `reference_reads_per_repeat` value is integer `1 // 4 == 0`; it is
rounding of one read over four repeats, not a zero-read result. Refusal and
resource lanes separately record zero reads.

The machine had 32 CPUs and launch load averages of 1.57, 2.18, and 2.44.
An unrelated boundary-check process used roughly one CPU during capture and
was left running; it is recorded as host contention. The profile does not
claim cross-machine timing equivalence.

Receipts are retained in [`performance/results/`](performance/results/). The
p50 comparison is
[`performance/results/performance-report.md`](performance/results/performance-report.md),
the machine and capture metadata are in
[`performance/results/capture-summary.json`](performance/results/capture-summary.json),
and [`performance/results/retained-files.json`](performance/results/retained-files.json)
hashes every retained result file except itself. The retained-files manifest
SHA256 is
`8f2cf58f08b33fd8b63b8af5953f1a3bba22a7f388736cd3f7d1758352fa4c1b`.
