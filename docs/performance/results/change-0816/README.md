# 0816 — finite-budget reads over a delayed caller source

This packet measures unchanged production at `c8ff2b9f65`. The reusable
`tools/perf-execution` harness gains an explicit cap on returned bytes and
bounded delay per nonempty successful `ReadAt` call. It is an in-memory source
simulation, not a network, disk, or page-cache benchmark. The full plan is
frozen before build and first capture.

The 72 cases cross CFB bulk reads and ordered OPC Parts, large/fresh,
mixed/fresh and large/primed sessions, widths 1/2/4/8, and three source
controls: uncapped/no-delay, 64 KiB cap/no-delay, and 64 KiB cap/250 µs delay.
The task floor is 64 KiB. Each sample has finite hierarchical budgets and
requires ordered exact payload hashes and released worker/I/O reservations.
Metadata setup and priming use the same provider but are outside the operation
clock; primed Parts are a zero-read cache control. Corpus generation and
verification are outside the clock. CPU time brackets a slightly wider scope,
and peak RSS covers the entire process.

Root runs all Cargo and workload commands serially. The two release binaries
separate ordinary timing from source-counter diagnostics. Qualification has
72 one-sample reports; native capture has six counterbalanced blocks, thirty
samples and three warmups (432 reports/12,960 samples); source observation
has two blocks with two samples and no warmup (144 reports/288 samples).
Total: 648 reports and 13,320 measured outputs. All children use CPUs 12–19.
Offline readers start only after all captures terminate.

Scaling ratios pair each width with width one in the same source model and
block. Source controls compare capped/local and delayed/capped at matching
widths. Nearest-rank process quantiles and median six-block ratios use 10,000
bootstrap resamples, seed 816816, endpoints 250/9749. Apparent Amdahl fractions
are diagnostic fits, not measured phases. There is no adoption gate or
production optimization claim; every tail, CPU, RSS and spread limitation is
retained. Historical timings are not pooled.

Original execution sequence (drivers refuse evidence overwrites):

```sh
python3 -B docs/performance/results/change-0816/quality.py
python3 -B docs/performance/results/change-0816/build.py
python3 -B docs/performance/results/change-0816/capture.py qualification
python3 -B docs/performance/results/change-0816/qualification_review.py --write
python3 -B docs/performance/results/change-0816/capture.py native
python3 -B docs/performance/results/change-0816/capture.py observer
```

The source and protocol reviews, raw reports, immutable numerical analyses,
independent audit, cleanup witness and final seal are retained in this packet.
The broader OLE2/OOXML goal remains active; ODF is deferred and iWork excluded.

Final offline replay (no Cargo or workload invocation):

```sh
python3 -B docs/performance/results/change-0816/qualification_review.py --check
python3 -B docs/performance/results/change-0816/validate.py --final --require-final-seal
```

Delayed large fresh CFB and Parts have paired width-eight speedups of 7.364×
and 5.674×. Mixed requests stay near 1×, while primed Parts have no source
reads inside the timed interval and slow down at higher requested widths.
These are configuration comparisons, not production improvements. Eight of
18 Amdahl fits are inadmissible. The first analysis and readers are retained
in `analysis-attempt-0/`; correcting the lower bootstrap endpoint changes no
numerical result, while the fit-admissibility flags change in 32 width rows.
No measured child was rerun or discarded. See the main report and independent
results review for complete limits.
