# Candidate release 02: union/cache/mask accounting analysis

This is a diagnostic comparison of the release-02 candidate capture with the
archived release-01 candidate capture. It separates exact work and accounting
fields from one-shot elapsed-time and RSS observations. It is not a
performance acceptance result.

## Scope and provenance

Release 02 completed all 854 expected rows with status `0`: 282 scalar rows,
396 value-evaluator rows, and 88 rows in each worksheet-adapter lane. The
scalar comparison has 282 rows, zero parity mismatches, and 67 rows with at
least one metric over the 5% review threshold. `automatic_acceptance` remains
false. The capture used three warmups, 15 samples, locked offline release
builds, and CPU 6. These are capture settings, not evidence of an idle host.

The source-closure receipts bind the copied candidate bytes rather than just a
Git `HEAD`:

| capture | closure files | closure SHA-256 | harness-input SHA-256 |
|---|---:|---|---|
| release 01 | 454 | `87bd96257b3680283afdc063280c91277f112854a4dac0b02894e961d70a963d` | `08c74145aa4be66f756ddadc190e99ca9fcdc6ca344495797bb72f647c508a0a` |
| release 02 | 455 | `2c679cebf213af53e44306ef0504ea813ed84ffbfe183d62153cbb5f1153c5b7` | `08c74145aa4be66f756ddadc190e99ca9fcdc6ca344495797bb72f647c508a0a` |

The release-02 `source-identity.json` and before/after closure records are
archive members in [`candidate-release-02.tar.gz`](diagnostics/candidate-release-02.tar.gz);
the release-01 equivalents are in
[`candidate-release-01.tar.gz`](diagnostics/candidate-release-01.tar.gz).

The archived closure comparison shows the scoped production entries changed
between these captures were `evaluation/value.rs` and
`evaluation/value/references.rs`; the closure also adds
`evaluation/value/tests.rs` and changes the array/reference integration test.
The candidate source places shape-mask and demand/condition-cache storage in
`crates/litchi-ods/src/codec/formula/evaluation/value.rs` (the
`ValueEvaluator` fields around line 1349), including their bounded
reservations in the mask helpers around line 3440 and cache helpers around
line 2497. Reference-set union/range construction is in
`crates/litchi-ods/src/codec/formula/evaluation/value/references.rs` around
line 573 in the copied candidate workspace recorded by the receipt.
This narrows the comparison to the intended evaluator changes, but does not
provide allocation-stack attribution.

The complete release-02 custody and row receipts are archive members
`run.json`, `compare.json`, and `review-triggers.json` in
[`candidate-release-02.tar.gz`](diagnostics/candidate-release-02.tar.gz).
The raw CSV hashes are:

| lane | archive member | SHA-256 |
|---|---|---|
| scalar | `scalar/raw.csv` | `9e63cf5561bfb60b98d535758cebb982a66383d59646716eb377735e99364d80` |
| value | `value/raw.csv` | `03b0b81fcbface7e1a896beaf19889ddd66d860f19c38a238d897c88c89072ce` |
| worksheet | `worksheet/raw.csv` | `fdedb7b01639404e810110500d87e84f0e08f0f432a8ce73611c00c7c99347fe` |
| worksheet instrumented | `worksheet-instrumented/raw.csv` | `04206cfaa1a5f2eba2769d033e37354f94b858fba246fc10d978bfd36edf3d32` |

Release 01 raw rows, custody receipts, and source closures are retained in
[`candidate-release-01.tar.gz`](diagnostics/candidate-release-01.tar.gz). The
release-02 value/worksheet executable SHA-256 is
`2be5b0c50845294cdada1651687d410d74f6c10c2d22e7b217cb3fa360480e09` and the
scalar executable SHA-256 is
`aaf55b305660c43269c02d3a81ad61c2acac9de6771c98f41fbd7ef8b37dba40`.

`requested_bytes` and `released_bytes` are allocator-instrumentation totals;
`peak_live_delta` is the instrumented high-water delta from the timer entry;
`memory_retained_used` is an execution-budget reservation sampled after the
result is produced, not a transient peak; and `max_rss_kib` is process-wide
maximum RSS. Value and worksheet `p50_ns` values are timed-batch values, with
the configured `repeat` inside the batch.

## Exact release-01 to release-02 accounting changes

Rows were joined by `(phase, case)` and compared across every non-timing and
non-RSS field. The only deterministic field changes were
`requested_bytes_{p50,max}`, `released_bytes_{p50,max}`, and
`peak_live_delta_{p50,max}`. There were no changes to work, allocation-call
counts, deallocation counts, retained-budget values, resolver counters,
success/refusal counts, checksums, failures, or statuses.

| lane and phase | rows | rows with deterministic accounting changes | changed fields |
|---|---:|---:|---|
| value / evaluate | 99 | 90 | requested, released, peak-live |
| value / parse-evaluate | 99 | 90 | requested, released, peak-live |
| worksheet / evaluate (both lanes) | 44 | 44 | requested, released, peak-live |
| worksheet / parse-evaluate (both lanes) | 44 | 44 | requested, released, peak-live |
| value setup/parse and worksheet setup/construct (both lanes) | 286 | 0 | none |
| scalar / all phases | 282 | 0 | none |

The changed bytes are visible in the following evaluate-phase scaling deltas.
Each delta is release 02 minus release 01 and is a per-timed-batch value; the
case repeat count changes with size.

| case family | size 1 | size 16 | size 256 | size 1,024 | size 4,096 |
|---|---:|---:|---:|---:|---:|
| reference-repeat, Δrequested / Δpeak-live | +2,048 / +16 | +512 / +16 | +128 / +16 | +32 / +16 | +16 / +16 |
| array-literal, Δrequested / Δpeak-live | +2,048 / +16 | +7,680 / +128 | +32,640 / +2,048 | +32,736 / +8,192 | +65,520 / +32,768 |
| sequence-and/or, Δrequested / Δpeak-live | +4,096 / +32 | +11,776 / +256 | +49,024 / +4,096 | +49,120 / +16,384 | +98,288 / +65,536 |

Representative rows show that the accounting deltas do not alter charged
work, allocation-call counts, or result-retained reservations:

| value evaluate case | requested bytes, release 01 → 02 | peak live, release 01 → 02 | work | alloc calls | retained bytes |
|---|---:|---:|---:|---:|---:|
| `scalar-number` | 78,848 → 80,896 | 616 → 632 | 384 | 512 | 0 |
| `reference-range-4096` | 722,040 → 722,056 | 721,576 → 721,592 | 8,199 | 10 | 360,448 |
| `reference-repeat-4096` | 2,163,176 → 2,163,192 | 263,496 → 263,512 | 16,382 | 12,304 | 0 |
| `array-literal-4096` | 2,391,864 → 2,457,384 | 1,376,440 → 1,409,208 | 23,471 | 29 | 360,448 |
| `matrix-lazy-inline-4096` | 3,277,080 → 3,342,600 | 1,737,304 → 1,770,072 | 72,629 | 8,252 | 360,448 |
| `sequence-and-4096` | 2,785,032 → 2,883,320 | 1,769,656 → 1,835,192 | 36,869 | 28 | 0 |

The phase selectivity and the shape of the deltas are consistent with the
scoped union/cache/mask implementation: only evaluation phases changed, and
the largest matrix paths acquire cell-scaled temporary accounting. This is a
source-and-receipt correlation, not proof that a particular cache or mask owns
each byte. `released_bytes` rises by the same amount as `requested_bytes`,
while `memory_retained_used` stays exact, so these rows show additional
transient accounting in this harness rather than a measured increase in
result-retained budget. A heap profile or component counters would be needed
to assign the bytes between union records, shape masks, and caches.

## One-shot elapsed time and RSS observations

Release 01 and release 02 are separate candidate windows, not paired A/B
samples. Their timing differences are therefore exploratory. For scale
context, the value evaluate p50 medians changed by +0.89% across 99 rows (range
−8.69% to +7.25%); parse had median −0.41% (−5.53% to +9.72%); and
parse-evaluate had median +1.19% (−7.45% to +10.79%). The uninstrumented
worksheet evaluate median changed +0.48% (−1.43% to +5.88%).

| lane/case | release 01 p50 ns | release 02 p50 ns | one-shot delta | RSS KiB 01 → 02 |
|---|---:|---:|---:|---:|
| value evaluate `reference-range-4096` | 772,973 | 755,963 | −2.20% | 4,428 → 4,444 |
| value evaluate `reference-repeat-4096` | 2,529,512 | 2,624,712 | +3.76% | 5,832 → 5,740 |
| value evaluate `reference-distinct-4096` | 2,598,451 | 2,684,622 | +3.32% | 5,996 → 5,728 |
| value evaluate `array-literal-4096` | 475,102 | 491,642 | +3.48% | 5,068 → 5,128 |
| value evaluate `sequence-and-4096` | 734,343 | 765,014 | +4.18% | 5,456 → 5,520 |
| value evaluate `matrix-lazy-inline-4096` | 1,909,529 | 1,896,018 | −0.71% | 6,956 → 6,824 |
| worksheet evaluate `adapter-repeated-row-4096` | 1,000 | 1,050 | +5.00% | 2,928 → 3,004 |
| worksheet evaluate `adapter-repeated-cell-4096` | 1,020 | 1,080 | +5.88% | 3,068 → 2,896 |

The scalar comparison against the retained baseline has zero parity mismatches
and 67 review rows (139 metric flags: 17 p50, 59 p95, 59 p99, and 4 RSS),
versus 78 rows (152 flags) in release 01. Release-02 p50 review rows are:

* evaluate: `bitwise-coerce-text` (+13.5%), `bitwise-error-shift` (+7.3%),
  `bitwise-lazy-selected` (+7.6%), `control-flat-256` (+5.1%),
  `logical-not-256` (+5.2%), `roman-error-arity` (+12.6%), and
  `roman-error-low` (+7.8%);
* parse: `arabic-input-256` (+5.1%), `bitwise-and-64` (+6.2%),
  `bitwise-rshift-4096` (+6.0%), `control-utf8-left-256` (+9.0%), and
  `failure-memory` (+6.3%);
* parse-evaluate: `bitwise-coerce-text` (+10.1%), `bitwise-error-shift`
  (+8.5%), `radix-decimal-small` (+5.7%), `roman-error-arity` (+6.5%), and
  `roman-error-low` (+7.5%).

Only four rows triggered the scalar RSS threshold: evaluate
`lazy-if-false-reference` (+5.0%), `logical-not-256` (+8.5%), and
`roman-concat-256` (+5.7%), plus parse `arabic-input-256` (+5.3%). Tail
percentiles and these one-shot timings remain review targets, not stable
latency or causality evidence.

## Separate owned-conversion scaling capture

The candidate-only `owned-release-01` archive
([`owned-release-01.tar.gz`](diagnostics/owned-release-01.tar.gz), member
`raw.csv`) contains 54 status-0 rows, three warmups, and 31 samples. Its raw
CSV SHA-256 is
`11c76916813d527051f2894956575780aacc769eadbc2f09c440b956cc08c1fc`;
the retained owned ELF is
`314b8a359dc20c67adcf67e595e8009405962272b1329a0270eeec538bde79a0`.
This is an absolute candidate-only capture under its own source manifest, not
an A/B comparison with release 01 or release 02.

The `own` phase prepares parse/evaluate before the timer and measures
`Evaluated::to_owned` plus owned-result inspection/drop. Its elapsed p50 is
already normalized by the configured repeat count. The other columns retain
the harness's batch or maximum semantics. The separate
`parse-evaluate-own` phase includes preparation and ownership and must not be
compared to `own` as if it covered the same operation.

| case | own p50 ns/op at sizes 1/16/256/4,096 | owned retained bytes at sizes 1/16/256/4,096 |
|---|---|---|
| inline array | 291 / 628 / 5,730 / 86,730 | 24 / 384 / 6,144 / 98,304 |
| duplicate reference list | 454 / 4,191 / 67,037 / 1,100,315 | 49 / 4,752 / 76,032 / 1,216,512 |
| 3-D reference list | 720 / 8,565 / 133,524 / 2,214,569 | 61 / 4,944 / 79,104 / 1,265,664 |

At size 4,096 the corresponding `own` allocation requests were 98,480 bytes
for the inline array, 1,216,688 for the duplicate list, and 1,265,840 for
the 3-D list; allocation-call p50 values were 3, 8,195, and 20,483. The
broader `parse-evaluate-own` p50 values were 687,023, 4,067,248, and
7,077,291 ns for those same cases. These numbers characterize the public
ownership boundary and its retained reservation; they do not establish a
speedup or identify a copy mechanism beyond the harness's typed/checksum
checks.

## Disposition

Release 02 preserves the captured result, resolver, work, status, and retained
budget fields while adding deterministic transient allocator accounting in
the evaluator phases. The largest observed deltas are bounded by the tested
matrix sizes, but this report does not generalize them beyond this corpus or
claim that a particular cache/mask/union object caused each byte. Elapsed and
RSS differences are one-shot observations, and scalar review triggers remain
open. Use paired AB/BA captures and component-level allocation counters or a
heap profile before making an acceptance, regression, or speedup claim.
