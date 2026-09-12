# 0534: reject the CFB physical role/FAT paired-prefix walk

The 0534 candidate replaced the private CFB physical-reconciliation loop's
per-sector FAT lookup with paired iteration over the role/FAT common prefix,
then reported a short FAT after the prefix. Source review found the exact
unclaimed-marker and missing-entry behavior, marker precedence, padding
compatibility, and allocation-free success path preserved. The candidate is
nevertheless rejected: all eight frozen primary XLS p50 rows are slower, so
no runtime speedup is retained. The runtime hunk is restored to the 0533
baseline; the two independent physical-layout contract tests remain.

The [frozen plan](../results/change-0534/plan.json) is bound to plan SHA-256
`4e1a13bdfa7f7c0268065d52def107f085b07cc0587d39685d9587de181d40dc` and the
priority `OLE2/OOXML active; ODF deferred until that goal completes; iWork
excluded`. The [final source manifest](../results/change-0534/final/source-manifest.json)
records `crates/litchi-cfb/src/file.rs` SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`: the
baseline runtime plus the retained tests. The measured baseline and candidate
source hashes were respectively
`bb1928970dc3d652c5091d521cad86c8873ab8352eee5d2eed18a6aa4dd4c20b` and
`a7dfb85e47de862f80350c335d4badc97df0b9cc68aa9ef1d2b8bd2309588e6f`.

## Matched evidence

The [matched comparison](../results/change-0534/comparison.json) contains
24,000 native samples and 720 allocator samples per stage: 48,000 native and
1,440 allocation samples overall. Its comparison status is `pass` for
identity and report validation; admission fails independently on the primary
latency rule, which requires at least 3% lower p50 in both repeats.

| Primary workflow | Repeat 1 p50 latency | Repeat 2 p50 latency |
| --- | ---: | ---: |
| `xls_source_backed_open` | 101,940 → 112,785 ns (+10.6386%) | 108,531 → 109,681 ns (+1.0596%) |
| `xls_source_backed_open_one_cell` | 107,005 → 115,941 ns (+8.3510%) | 112,321 → 115,130 ns (+2.5009%) |
| `xls_owned_source_open` | 96,970 → 103,891 ns (+7.1373%) | 100,240 → 105,180 ns (+4.9282%) |
| `xls_owned_source_open_one_cell` | 99,220 → 105,830 ns (+6.6620%) | 102,955 → 109,195 ns (+6.0609%) |

The separate allocator guard passes its required allocation-call, allocated-
byte, and incremental-region-peak checks without material growth. Its
instrumented elapsed time remains outside native latency evidence.

The [profile comparison](../results/change-0534/profile-comparison.json)
contains eight constructor children per stage: 16 children, 80 timed
constructor dumps, and 12 CFB setup dumps overall. It reports physical
reconciliation self Ir falling `16.6648257%` in both XLS-owned repeats and
`16.6648166%` in both CFB few-large repeats. XLS-owned constructor Ir falls
`2.8968%` in repeat 1 and `2.9363%` in repeat 2. These are per-workload
Callgrind diagnostics only and do not override the native regression or
establish a removable end-to-end cost.

All 153 recorded review flags remain in the bundle: 86 matched over-5%
adverse flags and 67 same-build over-5% variations. No row is discarded and
no cause or stable-tail claim is inferred. The [assembly analysis](../results/change-0534/assembly-analysis.json)
and raw vectors remain custody evidence for the rejected candidate; they do
not authorize retention.

## Restored source and remaining scope

The final source keeps `physical_layout_checks_every_role_and_fat_marker` and
`physical_layout_preserves_prefix_order_and_padding_contract`. Both call the
unchanged private reconciliation method and are independent of the rejected
paired-iterator implementation, so they remain valid baseline regression
coverage. No public API, dependency, unsafe code, validation policy,
allocation policy, or OLE2/OOXML ownership boundary changes are retained.

Retained candidate receipts record 14 checks and 4,382 candidate-stage test
executions. The restored source also passed all 14 checks and 4,382 test
executions, as recorded in the [final quality summary](../results/change-0534/quality-summary.json).
Both source states total 8,764 test executions. Cleanup and full
[post-cleanup verification](../results/change-0534/verification.json) passed;
the [evidence inventory](../results/change-0534/SHA256SUMS) is sealed. All
118 successful receipts form a verified serial timeline. This warm synthetic in-memory campaign makes no
cold/range, physical-provider, native-Office-producer,
scaling, broad CRUD, or ODF claim. OLE2/OOXML remains active; ODF stays
deferred and iWork remains excluded.
