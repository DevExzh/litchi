# Candidate release 01: array/reference evaluator diagnostic analysis

This is a diagnostic analysis of one candidate release capture. It compares
the established scalar API with the retained scalar baseline and records
absolute value-evaluator and worksheet-adapter measurements. It is not a
performance acceptance result: the capture uses one shared-host window, and
the comparison receipt leaves automatic acceptance disabled.

The candidate completed every child operation:

* scalar: 94 cases × 3 phases = 282 rows;
* value evaluator: 99 cases × 4 phases = 396 rows;
* worksheet adapter: 22 cases × 4 phases = 88 rows in each uninstrumented and
  instrumented lane.

All 854 rows have status `0`, expected result/checksum parity within the
candidate corpus, and no missing rows. The scalar comparison has 282 rows,
zero parity mismatches, and 78 rows with at least one metric above the 5%
review threshold. The complete machine-readable receipts are
[run.json (archive member `run.json`)](diagnostics/candidate-release-01.tar.gz),
[compare.json (archive member `compare.json`)](diagnostics/candidate-release-01.tar.gz),
and
[review-triggers.json (archive member `review-triggers.json`)](diagnostics/candidate-release-01.tar.gz).

## Measurement boundary and custody

The release was built in the isolated copied workspace recorded by
[source-identity.json (archive member `source-identity.json`)](diagnostics/candidate-release-01.tar.gz).
The gate-defined production closure contained 454 files and had SHA-256
`87bd96257b3680283afdc063280c91277f112854a4dac0b02894e961d70a963d` before
and after capture. The scalar candidate ELF is
`d13e23277c92c4de31236a4357dbc37220aa380991b1fa096b002231ad5feaa6` and the
value/worksheet ELF is
`ddfaa1a9cf611abc9e7ae465351647a211b8b48c6072a3defc1201bf9e51a905`.
The retained scalar baseline ELF is
`b68d801f62a1ec6e4accdd090940dda46e58d97c7818b91e4dbcc38ca8cdf682`.
The build used locked offline release mode, CPU 6 affinity, three warmups,
and 15 measured iterations. These are custody facts, not evidence that the
host was otherwise idle.

The candidate raw files are retained at
[scalar/raw.csv (archive member `scalar/raw.csv`)](diagnostics/candidate-release-01.tar.gz)
(`779bafabe4dfe9702cc7852777c845d2197e072af99cb9d25280545e7c5ed28f`),
[value/raw.csv (archive member `value/raw.csv`)](diagnostics/candidate-release-01.tar.gz)
(`5bf84fb09e204a46996c81b2eef1e1e5dc65ac3cc0a774eb202aad02da00f21b`),
[worksheet/raw.csv (archive member `worksheet/raw.csv`)](diagnostics/candidate-release-01.tar.gz)
(`2c404ddcea81a6445c3e922e5f45c7f5772fbd83a929c24013244d9ea7811803`), and
[worksheet-instrumented/raw.csv (archive member `worksheet-instrumented/raw.csv`)](diagnostics/candidate-release-01.tar.gz)
(`5c88795aa7e090c8694dad483e43aba0ce80dc93eb3a8cf81dc2a448403d0e6d`).
The compared retained baseline raw file is
[baseline/raw.csv](baseline/raw.csv), SHA-256
`22ed898910a76723d81d5e93fe5f40a8d9f8f20b5a2e37166ebfd01c562f54e0`.

`p50_ns`, `p95_ns`, and `p99_ns` are child-process timings for the harness
operation. A case's `repeat` count is inside the timed batch, so a p50 is a
batch time; dividing by `repeat` is only a descriptive per-operation
normalization. `work_used` and `memory_retained_used` are execution-budget
measurements. The latter is a retained reservation sampled after evaluation,
not a transient allocator peak. `alloc_calls`, `requested_bytes`, and
`peak_live_delta` are allocator instrumentation, while `max_rss_kib` is the
process maximum RSS. Resolver read and distinct-key counters are available in
the instrumented worksheet lane; their atomics and mutex bookkeeping add
overhead, so that lane is not a clean timing comparison with the uninstrumented
lane. No package publication or file I/O is included.

## Value evaluator scaling

The following are candidate-only absolute values from the evaluate phase. Each
row is one corpus case; `reads/distinct` is the resolver counter pair, `req`
is `requested_bytes_p50`, and `ret` is `memory_retained_used_p50` in bytes.
The array cases return a 64×64 matrix at size 4,096. The reference-range case
reads a range; reference-repeat reuses one key, while reference-distinct uses
different keys. This distinguishes provider activity from output size; it is
not a claim about all formulas or providers.

| case | repeat | p50 batch ns | work | reads/distinct | alloc calls | req | ret | RSS KiB |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| array-literal-1 | 128 | 83,901 | 640 | 0/0 | 768 | 107,520 | 176 | 2,976 |
| array-literal-16 | 32 | 84,361 | 1,824 | 0/0 | 416 | 292,608 | 1,408 | 3,188 |
| array-literal-256 | 8 | 182,061 | 9,392 | 0/0 | 168 | 1,194,432 | 22,528 | 3,204 |
| array-literal-1024 | 2 | 237,011 | 10,078 | 0/0 | 50 | 1,195,632 | 90,112 | 3,612 |
| array-literal-4096 | 1 | 475,102 | 23,471 | 0/0 | 29 | 2,391,864 | 360,448 | 5,068 |
| array-arithmetic-1 | 128 | 145,070 | 1,280 | 0/0 | 1,024 | 146,432 | 176 | 2,968 |
| array-arithmetic-16 | 32 | 142,090 | 2,464 | 0/0 | 448 | 337,664 | 1,408 | 2,976 |
| array-arithmetic-256 | 8 | 401,752 | 11,472 | 0/0 | 176 | 1,374,656 | 22,528 | 3,440 |
| array-arithmetic-1024 | 2 | 453,812 | 12,134 | 0/0 | 52 | 1,375,856 | 90,112 | 3,736 |
| array-arithmetic-4096 | 1 | 918,444 | 27,571 | 0/0 | 30 | 2,752,312 | 360,448 | 5,828 |
| reference-range-1 | 128 | 195,940 | 1,152 | 128/1 | 1,280 | 191,488 | 176 | 3,068 |
| reference-range-16 | 32 | 125,880 | 1,248 | 512/16 | 320 | 126,720 | 1,408 | 3,192 |
| reference-range-256 | 8 | 339,502 | 4,152 | 2,048/256 | 80 | 369,600 | 22,528 | 3,328 |
| reference-range-1024 | 2 | 351,251 | 4,110 | 2,048/1,024 | 20 | 362,736 | 90,112 | 3,684 |
| reference-range-4096 | 1 | 772,973 | 8,199 | 4,096/4,096 | 10 | 722,040 | 360,448 | 4,428 |
| reference-repeat-1 | 128 | 108,100 | 256 | 128/1 | 896 | 130,048 | 0 | 3,072 |
| reference-repeat-16 | 32 | 343,521 | 1,984 | 512/1 | 1,792 | 285,952 | 0 | 3,136 |
| reference-repeat-256 | 8 | 1,278,686 | 8,176 | 2,048/1 | 6,240 | 1,085,248 | 0 | 3,260 |
| reference-repeat-1024 | 2 | 1,283,996 | 8,188 | 2,048/1 | 6,172 | 1,082,320 | 0 | 3,984 |
| reference-repeat-4096 | 1 | 2,529,512 | 16,382 | 4,096/1 | 12,304 | 2,163,176 | 0 | 5,832 |
| reference-distinct-1 | 128 | 107,470 | 256 | 128/1 | 896 | 130,048 | 0 | 3,128 |
| reference-distinct-16 | 32 | 352,862 | 1,984 | 512/16 | 1,792 | 285,952 | 0 | 2,928 |
| reference-distinct-256 | 8 | 1,281,486 | 8,176 | 2,048/256 | 6,240 | 1,085,248 | 0 | 3,196 |
| reference-distinct-1024 | 2 | 1,290,616 | 8,188 | 2,048/1,024 | 6,172 | 1,082,320 | 0 | 3,992 |
| reference-distinct-4096 | 1 | 2,598,451 | 16,382 | 4,096/4,096 | 12,304 | 2,163,176 | 0 | 5,996 |

The successful array paths reserve 176 bytes at size 1 and 360,448 bytes at
size 4,096. Their work and requested bytes rise with the matrix size in this
corpus. Range reads and distinct keys likewise rise to 4,096 at the largest
case. Repeated and distinct references have nearly the same one-capture p50 at
4,096 (2.530 ms versus 2.598 ms), while the distinct-key counter separates
their provider behavior. The scalar reference paths also show substantially
more allocation activity than the range path at 256–4,096; the receipt gives
instrumented counts, but this capture does not establish why or whether the
pattern generalizes.

The value corpus also includes empty/error/limit/cancellation cases. All 396
candidate rows completed with their expected success or refusal status. A
large reservation or RSS value on an intentional refusal is therefore a
diagnostic of that bounded path, not a successful formula result.

## Worksheet adapter scaling

The uninstrumented lane is the timing reference and the instrumented lane
supplies read counters. Values below are evaluate-phase p50 batch times;
`work/op` and `reads/op` divide the batch counters by `repeat`. RSS is shown as
uninstrumented/instrumented KiB.

| case | size | p50 ns (plain/instrumented) | work/op | reads/op | RSS KiB (plain/instr.) |
|---|---:|---:|---:|---:|---:|
| repeated-row | 1 | 120,060 / 122,111 | 24 | 1 | 2,976 / 2,972 |
| repeated-row | 16 | 30,140 / 30,670 | 24 | 1 | 3,056 / 2,920 |
| repeated-row | 256 | 7,680 / 7,790 | 24 | 1 | 2,956 / 2,956 |
| repeated-row | 4,096 | 1,000 / 1,040 | 24 | 1 | 2,928 / 2,948 |
| repeated-cell | 1 | 122,151 / 123,811 | 24 | 1 | 3,184 / 2,972 |
| repeated-cell | 16 | 30,370 / 31,020 | 24 | 1 | 3,000 / 2,972 |
| repeated-cell | 256 | 7,670 / 7,910 | 24 | 1 | 3,212 / 2,924 |
| repeated-cell | 4,096 | 1,020 / 1,080 | 24 | 1 | 3,068 / 2,940 |
| background | 1 | 120,420 / 125,481 | 25 | 1 | 2,948 / 2,932 |
| background | 16 | 31,320 / 31,400 | 28 | 1 | 2,944 / 3,224 |
| background | 256 | 7,920 / 8,320 | 32 | 1 | 2,972 / 2,936 |
| background | 4,096 | 1,120 / 1,170 | 36 | 1 | 3,612 / 3,648 |

The repeated-row and repeated-cell fixtures keep one physical read and 24
work units per accessed cell across these sizes. Adding untouched background
content raises observed work from 25 to 36 units per operation and reaches
3,612/3,648 KiB RSS at size 4,096. That is evidence about this adapter fixture
and its index traversal, not a general locality or complexity proof. Resolver
counter overhead is why the instrumented p50 is reported separately.

## Scalar comparison and repeated review

The one-capture scalar comparison produced 78 unique trigger rows and 152
metric flags at the 5% review threshold. Counts by phase and metric were:

| phase | p50 flags | max RSS flags | p95 and p99 flags (same cases) |
|---|---|---|---|
| evaluate | `logical-ifna-4096`, `roman-error-arity`, `roman-error-text` | `control-coerce-1024`, `radix-decimal-max`, `roman-concat-256` | `arabic-input-256`, `bitwise-error-shift`, `control-escaped-256`, `failure-array`, `logical-and-1024`, `logical-and-64`, `logical-ifna-4096`, `logical-or-256`, `radix-fraction-error`, `roman-499-format-2`, `roman-error-arity`, `roman-error-text`, `roman-format-logical-true`, `text-if-concat` |
| parse | `arabic-input-256`, `bitwise-or-256`, `bitwise-rshift-4096`, `control-utf8-left-256`, `control-utf8-left-64`, `logical-and-64`, `logical-if-1024`, `logical-iferror-4096`, `logical-ifna-4096`, `radix-decimal-max`, `roman-499-format-1`, `roman-error-low` | `arabic-input-1024`, `arabic-input-256`, `roman-998-format-3` | `arabic-lowercase`, `bitwise-coerce-text`, `bitwise-lazy-selected`, `bitwise-lshift-4096`, `bitwise-or-256`, `bitwise-rshift-4096`, `control-coerce-64`, `control-escaped-64`, `control-flat-64`, `control-utf8-left-256`, `control-utf8-left-64`, `failure-array`, `failure-cancelled`, `lazy-if-false-reference`, `logical-not-256`, `logical-or-256`, `logical-true`, `logical-xor-4096`, `radix-base-max`, `roman-3888-format-0`, `roman-3888-format-1`, `roman-3888-format-2`, `roman-499-format-0`, `roman-998-format-1`, `roman-998-format-3`, `roman-bound-max`, `roman-error-arity`, `roman-error-low`, `roman-refusal-cancelled`, `roman-refusal-memory`, `roman-refusal-work`, `roman-truncate`, `roman-zero`, `text-if-concat`, `text-if-escaped` |
| parse-evaluate | `roman-error-arity` | `control-escaped-256`, `roman-refusal-cancelled` | `arabic-error-invalid`, `bitwise-coerce-text`, `bitwise-error-shift`, `control-utf8-left-256`, `control-utf8-left-4096`, `failure-cancelled`, `failure-reference`, `logical-false`, `logical-true`, `radix-base-max`, `roman-499-format-2`, `roman-499-format-4`, `roman-error-low`, `roman-refusal-stack`, `text-if-escaped` |

The p95/p99 entries above are one-capture tail observations. They are retained
for review but are not treated as stable tail latency evidence. The scalar
comparison did not flag allocator or retained-budget fields; all parity fields
were exact. The complete row-level deltas remain in
[review-triggers.json (archive member `review-triggers.json`)](diagnostics/candidate-release-01.tar.gz).

Three later paired scalar windows used the same inputs, harness, and retained
ELFs, with serial orders baseline→candidate, candidate→baseline, and
baseline→candidate on CPU 6. Their compare receipts are
[pair 1 (archive member `pair-1-compare.json`)](diagnostics/scalar-paired-01.tar.gz),
[pair 2 (archive member `pair-2-compare.json`)](diagnostics/scalar-paired-01.tar.gz),
and [pair 3 (archive member `pair-3-compare.json`)](diagnostics/scalar-paired-01.tar.gz);
the custody record is
[receipt.json (archive member `receipt.json`)](diagnostics/scalar-paired-01.tar.gz).
Each pair had 282 rows and zero parity mismatches. Trigger-row counts were
103, 99, and 96; the metric/case trigger sets contained 149, 147, and 137
entries. Their intersection contains 22 metric/case flags:

| phase | case | metric | candidate delta range across three pairs |
|---|---|---|---:|
| evaluate | bitwise-lazy-selected | p99 ns | +6.61% to +8.59% |
| evaluate | logical-and-64 | p95 ns | +9.64% to +18.99% |
| evaluate | logical-not-256 | max RSS | +9.64% to +11.42% |
| evaluate | roman-concat-256 | max RSS | +8.94% to +14.42% |
| evaluate | roman-error-arity | p50 ns | +10.77% to +11.92% |
| evaluate | roman-error-arity | p95 ns | +6.50% to +9.79% |
| parse | arabic-input-256 | p50 ns | +5.17% to +5.70% |
| parse | arabic-input-256 | p99 ns | +5.19% to +158.89% |
| parse | bitwise-error-shift | p99 ns | +12.48% to +48.26% |
| parse | bitwise-or-256 | max RSS | +9.01% to +14.08% |
| parse | control-utf8-left-256 | p50 ns | +7.41% to +9.37% |
| parse | control-utf8-left-256 | p95 ns | +8.09% to +9.00% |
| parse | control-utf8-left-64 | p99 ns | +6.90% to +47.74% |
| parse | lazy-if-false-reference | p95 ns | +6.08% to +9.42% |
| parse | logical-ifna-4096 | p95 ns | +5.71% to +11.61% |
| parse | radix-decimal-small | p99 ns | +23.80% to +54.85% |
| parse | radix-fraction-error | p99 ns | +7.74% to +63.72% |
| parse | roman-998-format-3 | p99 ns | +11.20% to +51.24% |
| parse-evaluate | control-coerce-1024 | max RSS | +5.58% to +13.37% |
| parse-evaluate | control-true | p99 ns | +5.11% to +25.09% |
| parse-evaluate | roman-499-format-3 | p99 ns | +5.13% to +10.04% |
| parse-evaluate | roman-error-text | p99 ns | +5.95% to +11.51% |

The strongest repeated central-tendency signals are the evaluate
`roman-error-arity` p50 (+10.77% to +11.92%) and parse
`control-utf8-left-256` p50 (+7.41% to +9.37%). Several initial p50 flags did
not persist: `logical-ifna-4096` fell to +1.63%–+3.18%, `roman-error-low`
fell to −0.17%–+1.33%, and the initial `roman-error-text` p50 increase was
not above 5% in every pair. The persistent tail and RSS rows are useful
follow-up targets, but their wide ranges include host and percentile noise;
the 158.89% `arabic-input-256` p99 result is one such outlier. No profile was
collected, so these receipts do not identify a code-level cause.

## Disposition

The candidate preserves scalar result parity and provides bounded absolute
evidence for matrix growth, reference read counts, worksheet background
traversal, budget reservations, allocator counters, and RSS. The value and
worksheet lanes have no baseline in this release, so they support workload
characterization only. The scalar A/B repetitions confirm a narrow set of
repeatable review signals, including two p50 tradeoffs, while disproving some
single-capture flags. Keep the candidate under review and do not describe it
as a speedup, a general locality guarantee, or an accepted regression-free
performance change.
