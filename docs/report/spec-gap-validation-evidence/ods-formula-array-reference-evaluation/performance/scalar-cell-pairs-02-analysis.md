# Scalar-cell allocation optimization: paired capture 02

This is a bounded diagnostic comparison of the preserved pre-optimization
value evaluator and the candidate scalar-cell reference path. It is evidence
for the direct scalar-reference allocation hypothesis, not acceptance of the
full performance goal.

## Scope and custody

The complete capture is retained in [`scalar-cell-pairs-02.tar.gz`](diagnostics/scalar-cell-pairs-02.tar.gz).
Its `run.json`, `pairs.csv`, `commands.txt`, and all child directories retain
the raw CSV, stdout, stderr, status, and `/usr/bin/time -v` files. The runner
completed 27 cases in three serial rounds, with one A/B pair per case in each
round: 162 child processes and 81 pairs, all process status `0`, with zero
deterministic mismatches.

The cases cover repeated and distinct scalar references at sizes 1, 16, 256,
and 4,096, plus range, matrix, text, empty, error, lazy, background, and
resource-refusal controls. Each child used three warmups and 31 measured
iterations. The frozen harness supplies the exact batch repeat:

| size | repeat inside each timed batch |
|---:|---:|
| 1 | 128 |
| 16 | 32 |
| 256 | 8 |
| 4,096 | 1 |

The `p50_ns` values below are raw timed batch values. They are not divided by
the repeat count. For the BA round, A and B are selected by child slot rather
than execution order, so every entry is consistently A (before) → B (after).

The before ELF is recorded by [`scalar-cell-before.tar.gz`](diagnostics/scalar-cell-before.tar.gz)
with SHA-256 `0416017dc236eb55acbd79b0a7efae7131924642ac9ba03aaef7763669b6e1da`
and source head `fca1f39d4`. The candidate ELF is recorded by
[`scalar-cell-after-01.tar.gz`](diagnostics/scalar-cell-after-01.tar.gz)
with SHA-256 `e3afee412defa04c25e79586e986eac89164316646f4bcb9536f2d7d80015187`.
Both source receipts contain 455-file source closures. The four frozen value
harness files were unchanged throughout the capture:

| harness file | SHA-256 |
|---|---|
| `Cargo.toml` | `44b731de34b54ee723e9611b1c32cfc2ee91a802b674494dc02a20fafa89819a` |
| `Cargo.lock` | `25e9c37f67ec26b73e760d8eb00a581cfa4ac7e1e425b8c42d7d1ce450878d0a` |
| `run.py` | `f2b8ef3e06fccd3988fe87b866577405709bde264b3807244ed5c8c760b7d6e4` |
| `src/main.rs` | `6a02a57d210550e143469dabb0e16295cc64e4af6f17ff7e1cee06ee244fc583` |

The runner pinned the serial window to CPU 6. CPU affinity does not establish
an idle host; elapsed and RSS values remain subject to shared-host noise.

## Allocation and accounting result

The allocator counters below are p50 values from the evaluate batch. The
allocation shape is identical for repeated and distinct references, although
the provider distinct-key counter differs as expected.

| cells | repeat | alloc calls A → B | requested bytes A → B | peak-live delta A → B | retained bytes A → B |
|---:|---:|---:|---:|---:|---:|
| 1 | 128 | 896 → 512 | 132,096 → 80,896 | 1,032 → 632 | 0 → 0 |
| 16 | 32 | 1,792 → 256 | 286,464 → 81,664 | 2,392 → 1,592 | 0 → 0 |
| 256 | 8 | 6,240 → 96 | 1,085,376 → 266,176 | 17,752 → 16,952 | 0 → 0 |
| 4,096 | 1 | 12,304 → 16 | 2,163,192 → 524,792 | 263,512 → 262,712 | 0 → 0 |

At 4,096 cells, both repeated and distinct chains therefore show 12,304 →
16 allocation calls and 2,163,192 → 524,792 requested bytes. The allocator
instrumentation reports a 99.87% reduction in calls and a 75.74% reduction in
requested bytes for this batch. `memory_retained_used` remains zero for these
scalar chains; it is an evaluator reservation counter, while RSS is a
process-level maximum and is not an allocator measurement.

The reference range is a useful shape control. At 4,096 cells it remains
10 → 10 allocation calls, 722,056 → 722,056 requested bytes, and
360,448 → 360,448 retained bytes. This is not a per-cell equivalent: the
range materializes one rectangular result.

## Timed scalar-reference lanes

| case | repeat | A → B p50 ns, round AB | A → B p50 ns, round BA | A → B p50 ns, round AB |
|---|---:|---:|---:|---:|
| `reference-repeat-1` | 128 | 108,310 → 60,680 (−43.98%) | 113,210 → 62,310 (−44.96%) | 112,050 → 60,840 (−45.70%) |
| `reference-repeat-16` | 32 | 365,891 → 171,130 (−53.23%) | 361,771 → 159,861 (−55.81%) | 347,511 → 160,360 (−53.85%) |
| `reference-repeat-256` | 8 | 1,364,434 → 559,552 (−58.99%) | 1,354,514 → 564,661 (−58.31%) | 1,339,904 → 553,631 (−58.68%) |
| `reference-repeat-4096` | 1 | 2,740,947 → 1,128,783 (−58.82%) | 2,796,078 → 1,111,523 (−60.25%) | 2,698,128 → 1,197,233 (−55.63%) |
| `reference-distinct-1` | 128 | 107,941 → 61,160 (−43.34%) | 111,400 → 61,810 (−44.52%) | 111,971 → 62,700 (−44.00%) |
| `reference-distinct-16` | 32 | 348,571 → 164,611 (−52.78%) | 355,290 → 160,270 (−54.89%) | 357,421 → 167,971 (−53.00%) |
| `reference-distinct-256` | 8 | 1,371,124 → 625,692 (−54.37%) | 1,348,303 → 553,892 (−58.92%) | 1,370,644 → 581,012 (−57.61%) |
| `reference-distinct-4096` | 1 | 2,746,238 → 1,287,423 (−53.12%) | 2,706,657 → 1,140,253 (−57.87%) | 2,757,368 → 1,172,744 (−57.47%) |

The repeated-reference p50 reduction is 53–60% across sizes 16, 256, and
4,096; the distinct-reference reduction is 53–58% there. These are observations from one
paired capture window and are not a general throughput claim.

## Result and control parity

Across all 81 pairs, the following fields matched between A and B: process and
row status, expected success/refusal counts, failure label, checksum, resolver
read and distinct-read counts, work counters, borrowed-text bytes, and copied
bytes. Representative 4,096-cell values were:

| case | reads / distinct | work | checksum | borrowed / copied bytes |
|---|---:|---:|---:|---:|
| `reference-repeat-4096` | 4,096 / 1 | 16,382 | 16,140,901,064,495,858,248 | 0 / 0 |
| `reference-distinct-4096` | 4,096 / 4,096 | 16,382 | 2,251,799,813,685,829 | 0 / 0 |
| `reference-range-4096` | 4,096 / 4,096 | 8,199 | 2,830,450,581,222,239,528 | 0 / 0 |

The only consistent positive p50 control outlier was the smallest range:

| control | repeat | A → B p50 delta, AB / BA / AB | accounting parity |
|---|---:|---:|---|
| `reference-range-1` | 128 | +8.67% / +11.57% / +9.78% | 1,280 alloc calls; 193,536 requested bytes; 1,152 work; 128 reads / 1 distinct |

No other control had a positive p50 delta above 5%; `reference-matrix` was the
largest remaining positive control at +3.86%, +3.24%, and +4.47%. The range-1
increase has unchanged read/work/allocation counters, so this capture cannot
identify its cause. A fixed dispatch or code-layout effect is one possible
source-level explanation, but it is only an inference; host timing variation
and no CPU profile remain live alternatives. Isolated RSS differences are
also not treated as stable regressions.

## Disposition

The candidate removes the observed temporary allocation shape for direct
scalar references while preserving result, resolver, work, checksum, typed
failure, and text borrow/copy behavior in this corpus. The measurement covers
the evaluator's timed batch after harness setup; it does not include building
the caller's expression or publishing a package. Allocation counters and
requested bytes are instrumentation evidence, not proof of a particular copy
or CPU cause.

This remains a diagnostic result with `automatic_acceptance` false. A second
independent paired window and a focused CPU profile would be needed to turn
the timing observations into a stable performance disposition, especially for
the range-1 control regression.

## Current candidate 03: variant-ordering follow-up

Capture 03 is the current candidate follow-up. It changes only the internal
variant ordering so `ScalarCell` is placed after the existing `Array` and
`Areas` variants. The complete receipt is [`scalar-cell-pairs-03.tar.gz`](diagnostics/scalar-cell-pairs-03.tar.gz);
the source/variant receipt is [`scalar-cell-layout-02.json`](diagnostics/scalar-cell-layout-02.json).
The candidate ELF SHA-256 is
`2ac378a8c07350afb26b7ab7f79e0d7af7c7100c0c52dea42e21e93a5c182a22`.

The same 27-case, CPU-6 AB/BA/AB capture completed 162 child processes and 81
pairs with status `0` and zero deterministic mismatches. The 4,096-cell scalar
allocation result is unchanged from capture 02 (12,304 → 16 calls and
2,163,192 → 524,792 requested bytes). Its p50 timing reductions remain:

| case | A → B p50 delta, AB / BA / AB |
|---|---:|
| `reference-repeat-4096` | −58.95% / −59.79% / −57.82% |
| `reference-distinct-4096` | −58.96% / −58.24% / −58.75% |

The range-1 control improved but still crosses the 5% review threshold in all
three rounds (+5.67% / +5.46% / +6.09%). `reference-matrix` was +4.52% /
+5.64% / +4.92%, and `matrix-lazy-inline-4096` was +2.42% / +5.32% /
+0.33%; these remain control review signals with unchanged accounting and
result parity. The ordering variant therefore does not close acceptance: the
residual positive control timing flags require further paired review, and no
CPU cause or general speedup is claimed.

The [whole-process counter diagnostic](diagnostics/scalar-cell-range-stat-01.json)
recorded a 0.53% instruction increase and 7.18% cycle increase for candidate01's
one-cell matrix range, with more branch misses. The reordered candidate's
corresponding diagnostic is retained in
[scalar-cell-layout-02.tar.gz](diagnostics/scalar-cell-layout-02.tar.gz).
These single sequential counter runs include setup and warmups; they support
further investigation but do not identify a causal function. The final reordered
source passes [all five ODS gates](../gates/scalar-cell-candidate-02.json).
