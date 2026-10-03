# 0886 Producer XLSX full-cell profile

Current-source profiling identifies MCE processing as the first area to
investigate in the generated producer XLSX full scan. Three native cycle
profiles and two separate Callgrind diagnostics pass semantic replay. Production
code is unchanged; this batch adopts no optimization and claims no speedup.

The retained release executable and source revision are exactly those used for
[0885](0885-producer-full-cell-release-baseline.md). Current HEAD at capture was
`205a8a34e6`; production and harness source still match `335a40350b`. The ELF
build ID in the perf capture matches the retained executable. The 0885 baseline
remains the latency reference; profiler elapsed values are not substituted.

## Native sampled instruction pointers

The [frozen plan](results/change-0886/plan.json) captures three sequential fresh
processes on CPU 12 with `cycles:u`, frequency 499 Hz, and 16 KiB DWARF stack
samples. Each process runs ten warmups and 2,000 measured full scans. All 6,000
measured scans pass the same archive, cell-count, ordered-value, source-version,
and producer-census checks as 0885.

This is a whole-process profile: launcher/startup, corpus generation, oracle
setup, source preparation, timed scanning, post-timer checks, destruction, and
reporting are included. It is not an exact profile of the benchmark timer.

The [self-cost reducer](results/change-0886/reduce.py) independently sums every
sampled IP's period and reconciles all per-symbol periods and sample counts
with the flat `perf report` output. The populations contain 4,141, 4,133, and
4,096 samples, with zero reported lost samples. Period weights, not raw sample
counts or rounded display percentages, determine the following shares.

| Disjoint sampled-IP symbol group | Period-weighted share across three processes |
| --- | ---: |
| MCE symbols | 26.73–27.78% |
| x14ac symbols | 5.69–6.36% |
| XML namespace resolver symbols | 6.87–10.16% |
| XML namespace reader symbols | 4.00–5.13% |
| Unresolved sampled IPs | 0.60–0.92% |

These are self-symbol groups, with explicit prefix rules in the reducer; they
do not attribute shared XML helpers to a caller. Other resolved symbols remain
in the [full analysis](results/change-0886/self-analysis.json). Unknown IPs stay
in the denominator and are not silently assigned to application or kernel work.
The period sums are sampling estimates, not an independent hardware counter
measurement of total workload cycles. Observed ranges are descriptive, not
confidence intervals.

DWARF unwinding is insufficient for caller or inclusive attribution: empty
stacks represent 87.19%, 86.75%, and 69.47% of weighted periods. The stack
exports are retained, but no caller-cost claim or flamegraph percentage is
derived from them. Direct-IP self-symbol reports remain independently usable.

## Separate guest-instruction diagnostic

The [Callgrind plan](results/change-0886/callgrind-plan.json) adds two fresh
processes, each with ten measured scans and no warmups. Collection starts and
stops at the real `producer_shape::run_xlsx_full_cell_scan` function symbol.
The scope includes the runner's oracle preparation and other timer-excluded
work. Its timed closure is inlined, so this binary cannot supply a separate
timer-only function marker.

Both diagnostics pass all twenty scan checks. Total guest instruction counts
are 755,248,408 and 754,020,187. Canonical function/file rows in the retained
annotations give:

| Inclusive function | Share of collected guest instructions |
| --- | ---: |
| MCE `process_markup_compatibility` | 44.10–44.19% |
| MCE `start` | 28.81–28.85% |
| x14ac `capture_inner` | 19.21–19.24% |

These rows overlap: `start` is inside MCE processing. They must not be summed,
combined with native sampled cycles, or interpreted as native time fractions.
The [reader](results/change-0886/analyze_callgrind.py) selects exact canonical
function/file rows and checks the runner against the raw instruction total.

Valgrind reports a `brk segment overflow` warning. The successful outputs are
retained as instrumented guest-execution diagnostics, not native allocation,
memory, or latency evidence. Forked child diagnostics echo inherited counts;
the final raw profile header matches the original parent PID and its totals
are read once. An early progress message mistakenly called the two totals
identical after reading the first run's echoed count; the retained analyses
and figures above use the two distinct parent totals.

## Validation, custody, and next experiment

The [independent review](results/change-0886/review.json) checks scope,
unresolved symbols, weighted self attribution, and diagnostic limitations.
Both reducers reproduce the retained analyses. The first self-reader attempt
rejected three unit-period launcher samples because one view named unresolved
symbols `[unknown]` and the other used raw addresses. Normalizing those
unresolved names fixed the reader; no workload capture was repeated.

Raw reports, commands, exporter receipts, text sample exports, Callgrind data,
and annotations are tracked. The approximately 200 MiB of original `perf.data`
files remain locally retained and ignored by Git, with hashes in each capture
receipt. They use the locally retained 0885 executable for symbolization.
Text reducers can be replayed into a new output path:

```sh
python3 -B docs/performance/results/change-0886/reduce.py --output /tmp/0886-self.json
python3 -B docs/performance/results/change-0886/analyze_callgrind.py --output /tmp/0886-ir.json
```

Within this scan scenario, the next bounded experiment should examine MCE
`start`, including temporary attribute storage and repeated name expansion.
Those mechanisms are hypotheses, not yet measured savings. x14ac's discarded
namespace resolutions remain a secondary candidate. Any change needs matched
current-source before/after measurements, ordinary and MCE/x14ac controls,
shared DOCX/PPTX regression coverage, and unchanged namespace, alternate-content,
malformed-input, error-precedence, resource-limit, and preservation behavior.

No source build or Git worktree was created for this batch. Temporary replay
directories and exploratory duplicate exports were cleaned; original profile
data and canonical evidence are retained. The broader non-iWork performance
program remains open.
