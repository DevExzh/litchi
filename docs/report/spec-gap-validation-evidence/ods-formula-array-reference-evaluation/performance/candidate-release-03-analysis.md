# Candidate release 03: release 02 comparison

This is a diagnostic comparison of the release-03 capture with the archived
release-02 capture. It separates exact result/accounting fields from timing and
RSS observations. It is not a performance acceptance result.

## Scope and provenance

Release 03 completed all 854 expected rows with process status `0`: 282 scalar
rows, 396 value-evaluator rows, and 88 rows in each worksheet-adapter lane.
Release 02 has the same row counts and status results. Both captures used
three warmups, 15 measured iterations, locked offline release builds, and
CPU 6. Each is a separate serial window on a shared host; CPU affinity does
not establish an idle machine.

The status-0 rows certify only that the retained harness protocol completed for
this corpus. A separate review has reproduced a reference-list-kind correctness
issue on the release-03 source snapshot; its correction is pending. This
capture therefore does not close correctness review.

The complete receipts are retained in [`candidate-release-03.tar.gz`](diagnostics/candidate-release-03.tar.gz)
(SHA-256 `dba63de7009e214afde682de7013589ba733ec92e85071209934ee85982fe9b8`)
and [`candidate-release-02.tar.gz`](diagnostics/candidate-release-02.tar.gz)
(SHA-256 `f65202f07193c4518a01444140810d308f26cbbcc41cd30118ef02e813c8f911`).
Their `run.json`, `source-identity.json`, closure records, `compare.json`, and
`review-triggers.json` members provide the custody and row-level evidence.

| capture | source closure files | canonical/workspace closure SHA-256 | canonical head | harness-input SHA-256 |
|---|---:|---|---|---|
| release 02 | 455 | `2c679cebf213af53e44306ef0504ea813ed84ffbfe183d62153cbb5f1153c5b7` | `2073bbd385a6379aeb8b2a25342401bf99aee244` | `08c74145aa4be66f756ddadc190e99ca9fcdc6ca344495797bb72f647c508a0a` |
| release 03 | 455 | `77f528bf3ec2f82599685a6cc1c3d9d40f05bccb1296d60d584f7ee34ca5df7b` | `64a9b591302280f5142b61dcb5e83b0e867f8375` | `08c74145aa4be66f756ddadc190e99ca9fcdc6ca344495797bb72f647c508a0a` |

The closure comparison found no added or removed source entries and exactly
three changed entries:

| source-closure entry | release 02 SHA-256 | release 03 SHA-256 |
|---|---|---|
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `e96b23bc211b1af51295a182a98bb63f2b9b367dc8e737926446c50b909f9196` | `7766e372d346aa9e8473776a40e00c5127e14289417ce124a27a1ca5a3892b4d` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `c76e3ddb2546b151a3189a578c5c08208e500c5264de4ba9fa92dff9ea0c464e` | `e351dcd1589c6c97dca115a1c665468ca35b9f25c57ab03226e79029c4ce3dd4` |
| `crates/litchi-ods/tests/ods_formula_array_reference_evaluation.rs` | `b4536e8f534ec3210d75575db81d01c078d21074ed19a3849fef210e54e6ce00` | `4a7e8ecba78f32237f6da394d0b6a26c1c5b870c000f9313c8cd4e9200f142e6` |

The release-03 scalar executable SHA-256 is
`283467bbe71717e7d1eaf590813e512adb18d8d5851f03b5f1269c9b527c198f`; its
value/worksheet executable SHA-256 is
`ff20e65b8f4326b7fe2c924313c98a1e113543047bfd28aa369986b3be28e0f7`.
The release-02 executables were
`aaf55b305660c43269c02d3a81ad61c2acac9de6771c98f41fbd7ef8b37dba40` (scalar)
and `2be5b0c50845294cdada1651687d410d74f6c10c2d22e7b217cb3fa360480e09`
(value/worksheet). The run receipts record unchanged executable hashes before
and after each capture.

The raw CSV members and hashes are:

| lane | release-02 member SHA-256 | release-03 member SHA-256 |
|---|---|---|
| scalar/raw.csv | `9e63cf5561bfb60b98d535758cebb982a66383d59646716eb377735e99364d80` | `9b623fda197578fe7152a835e0c773d781fcf53f4632a195cba4a79889679a04` |
| value/raw.csv | `03b0b81fcbface7e1a896beaf19889ddd66d860f19c38a238d897c88c89072ce` | `7ae94def4b0ffb1e7cca99376e3c7d7807965ba0e7489918b3a489bd5968b35a` |
| worksheet/raw.csv | `fdedb7b01639404e810110500d87e84f0e08f0f432a8ce73611c00c7c99347fe` | `46dfc6b7356e448d8e10513c46aa62dcafacbb21327b932dab28e3893171a86b` |
| worksheet-instrumented/raw.csv | `04206cfaa1a5f2eba2769d033e37354f94b858fba246fc10d978bfd36edf3d32` | `4e6a2977565bbff1b370f63f79e9a5f7ce4f9a7e428aef24be1981c757199d9e` |

## Exact release-02 to release-03 parity

Rows were joined by `(phase, case)`. Every lane retained the same key set and
row count. Comparing every shared field except `mean_ns`, `p50_ns`, `p95_ns`,
`p99_ns`, `max_rss_kib`, `started_at`, and `finished_at` found zero differences:

| lane | rows 02 → 03 | release-03 process-status-0 rows | deterministic field changes |
|---|---:|---:|---:|
| scalar | 282 → 282 | 282 | 0 |
| value | 396 → 396 | 396 | 0 |
| worksheet | 88 → 88 | 88 | 0 |
| worksheet-instrumented | 88 → 88 | 88 | 0 |

This exact comparison covers allocator calls, requested/released bytes, live
and peak-live counters, work, retained execution-budget reservations, adapter
reservations, resolver reads/distinct reads/geometry calls, borrowed/copied
byte counters, pointer checks, success/refusal counts, checksums, failure
labels, expected-result fields, and statuses. The non-`none` failure labels
are expected refusal cases; all child processes still exited `0`.

| lane | release-03 `failure` values (row counts) |
|---|---|
| scalar | `none` 260; `resource-work` 4; `resource-memory` 4; `resource-objects` 4; `cancelled` 4; `unsupported-reference` 2; `unsupported-array` 2; `unsupported-name` 2 |
| value | `none` 378; `resource-objects` 12; `resource-work` 2; `resource-memory` 2; `cancelled` 2 |
| worksheet and instrumented | `none` 88 in each lane |

Representative 4,096-cell rows show the exact accounting and read parity. The
elapsed values below are the raw row `p50_ns` batch values, not normalized
single-cell timings.

| value evaluate case | p50 ns 02 → 03 | alloc calls | work | retained bytes | resolver reads | RSS KiB 02 → 03 |
|---|---:|---:|---:|---:|---:|---:|
| `reference-range-4096` | 755,963 → 756,493 (+0.07%) | 10 → 10 | 8,199 → 8,199 | 360,448 → 360,448 | 4,096 → 4,096 | 4,444 → 4,388 |
| `reference-repeat-4096` | 2,624,712 → 2,762,012 (+5.23%) | 12,304 → 12,304 | 16,382 → 16,382 | 0 → 0 | 4,096 → 4,096 | 5,740 → 5,988 |
| `reference-distinct-4096` | 2,684,622 → 2,717,372 (+1.22%) | 12,304 → 12,304 | 16,382 → 16,382 | 0 → 0 | 4,096 → 4,096 | 5,728 → 5,748 |
| `matrix-lazy-inline-4096` | 1,896,018 → 1,934,098 (+2.01%) | 8,252 → 8,252 | 72,629 → 72,629 | 360,448 → 360,448 | 0 → 0 | 6,824 → 6,940 |
| `matrix-lazy-aggregate-4096` | 1,676,577 → 1,686,987 (+0.62%) | 4,161 → 4,161 | 49,176 → 49,176 | 360,448 → 360,448 | 4,096 → 4,096 | 5,768 → 5,776 |

The exact parity establishes that this corpus did not change its observed
work, allocation, reservation, resolver, or result behavior between the two
source snapshots. It does not assign a timing difference to a particular
changed function.

## One-window elapsed time and RSS

The following are medians of the row `p50_ns` values and of the row
`max_rss_kib` values within each lane/phase. They summarize two independent
windows; they are not paired A/B estimates.

| lane / phase | p50 ns 02 → 03 | RSS KiB 02 → 03 |
|---|---:|---:|
| scalar / parse | 24,545 → 24,580 (+0.14%) | 2,620 → 2,646 (+0.99%) |
| scalar / evaluate | 189,920 → 190,581 (+0.35%) | 2,620 → 2,620 (+0.00%) |
| scalar / parse-evaluate | 210,161 → 208,686 (−0.70%) | 2,620 → 2,620 (+0.00%) |
| value / setup | 36,950 → 36,730 (−0.60%) | 2,656 → 2,648 (−0.30%) |
| value / parse | 37,840 → 37,560 (−0.74%) | 2,720 → 2,720 (+0.00%) |
| value / evaluate | 221,141 → 221,101 (−0.02%) | 2,988 → 3,108 (+4.02%) |
| value / parse-evaluate | 284,142 → 286,351 (+0.78%) | 3,016 → 3,132 (+3.85%) |
| worksheet / setup | 22,325 → 21,990 (−1.50%) | 2,648 → 2,672 (+0.91%) |
| worksheet / construct | 24,105 → 24,370 (+1.10%) | 2,708 → 2,658 (−1.85%) |
| worksheet / evaluate | 30,735 → 30,150 (−1.90%) | 2,972 → 3,092 (+4.04%) |
| worksheet / parse-evaluate | 38,636 → 39,765 (+2.92%) | 2,946 → 3,004 (+1.97%) |
| worksheet-instrumented / setup | 22,750 → 22,076 (−2.96%) | 2,678 → 2,672 (−0.22%) |
| worksheet-instrumented / construct | 24,336 → 23,985 (−1.44%) | 2,654 → 2,666 (+0.45%) |
| worksheet-instrumented / evaluate | 32,655 → 31,885 (−2.36%) | 2,952 → 2,972 (+0.68%) |
| worksheet-instrumented / parse-evaluate | 40,530 → 39,600 (−2.30%) | 2,942 → 2,994 (+1.77%) |

Across individual rows, release 03 versus release 02 had p50 deltas above
`+5%` / below `−5%` in scalar `3/5`, value `14/9`, worksheet `0/2`, and
instrumented worksheet `0/5` rows. RSS deltas exceeded `5%` in absolute value
in `21`, `83`, `20`, and `18` rows respectively. Those counts reflect the
single-window comparison and small KiB denominators; they are review signals,
not stable regressions or improvements.

The release-03 scalar comparison against the retained baseline is separately
recorded in the release-03 `compare.json` and `review-triggers.json` archive
members. It has zero parity mismatches and 65 review rows at the 5% threshold:
18 evaluate, 26 parse, and 21 parse-evaluate rows. The trigger fields count is
57 `p95_ns`, 57 `p99_ns`, 14 `p50_ns`, and 5 `max_rss_kib`. The preceding
release-02 comparison had 67 review rows, so the count difference is not an
acceptance signal.

## Corpus coverage boundary

The value corpus covers direct references, rectangular references, repeated and
distinct reference lists, the scalar skipped-reference `reference-lazy` case,
`matrix-lazy-inline-*`, and `matrix-lazy-aggregate-*` through size 4,096. The
aggregate lazy case evaluates `AND` over a reference range and therefore gives
one reference-bearing lazy branch with read/work scaling. `array-iferror` and
`array-ifna` exercise evaluated array operands, but their operands are inline
arrays rather than references. There is no reference-bearing `IFERROR` or
`IFNA` case, and no case that isolates every new kind/geometry planner path
with those handlers. Consequently, the unchanged counters above do not
measure the cost of those absent combinations. The 4,096-cell work and read
rows are separate bounded scaling evidence for the cases that are present.

## Disposition

Release 03 has exact result and accounting parity with release 02 over all
854 rows. Its elapsed and RSS differences are observations from separate
single windows, including a representative `+5.23%` value-evaluate timing
change for repeated references, with no corresponding accounting change.
The capture remains diagnostic (`automatic_acceptance: false`) and the known
correctness issue keeps acceptance open; paired AB/BA/repeated captures using
the same corpus would be required for a stable performance disposition after
that issue is resolved.
