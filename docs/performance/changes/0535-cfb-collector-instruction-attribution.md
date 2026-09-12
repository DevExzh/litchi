# Change 0535: CFB collector instruction-address attribution baseline

**Date:** 2026-09-12
**Status:** Diagnostic baseline; no runtime change retained
**Performance claim:** none

## Decision

This batch records instruction-address and generated-code evidence for the
existing CFB exact chain collector. It makes no production or harness change,
does not compare a candidate against a control, and establishes no speedup.
The source is exactly the final 0534 source, including its retained physical
layout contract tests. The CFB few-large repeat-2 p50 is 10.0959% higher than
repeat 1, which is a same-build repeat drift observation rather than a
candidate regression or a causal explanation.

The evidence is useful for locating the remaining cost in
`SectorChainScratch::collect_exact` and its checked visited map. It does not
authorize removing chain validation, ownership claiming, physical
reconciliation, allocation checks, or source/provider freshness boundaries.
No optimization is accepted from this diagnostic. Full postcleanup verification
passed; both owned scratch paths were removed and the evidence is sealed.

## Source and binary custody

The build is pinned to revision
`bb5bfaa7be4fdffeae3fddaee3ed1266c3417520`. The source manifest contains
8,585 entries and has SHA-256
`9d6a53738f299107582b71e5ddab03922b9646771df4d77b639b3fdc7fd6ff75`.
The reviewed `crates/litchi-cfb/src/file.rs` entry is
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`,
matching the final 0534 source manifest. The normal baseline executable is
60,139,544 bytes with SHA-256
`caa376f8313f4c2342e20ca4d3ee5126b06805206581fe276dbf5a76ab94028a`.

The [baseline source manifest](../results/change-0535/baseline/source-manifest.json),
[binary record](../results/change-0535/baseline/binary-normal.json),
[symbols receipt](../results/change-0535/baseline/symbols.receipt.json), and
[assembly index](../results/change-0535/baseline/assembly-index.json) bind the
source and generated code. The assembly index records nine successful
function captures: `CheckedBitSet::try_with_capacity` (236 bytes),
`CheckedBitSet::insert` (223 bytes), `SectorChainScratch::collect_exact`
(1,436 bytes), and six `SectorChainScratch` drop variants (62 bytes each).
The normal build, symbol extraction, nine assembly captures, four native
captures, and four profile captures make 19 baseline receipts. Every receipt
has execution stage `baseline`, exit code zero, and the same source manifest
hash. The symbol, assembly, native, and profile receipts also carry the same
binary hash; the normal build receipt's separate binary field is null and is
bound by `binary-normal.json`.

The [frozen plan](../results/change-0535/plan.json) has SHA-256
`6f1fa017c41d2e9e48043795509479dc28e566508c4eb7545c847fece4611d34`.
The [analysis record](../results/change-0535/analysis.json) has SHA-256
`fc5108224b9eb7c619c48422c240785147b7dcba1b697edd3e66324b51ebb155` and
reports the native evidence schema as `pass`; its scope remains diagnostic
only. Profile clocks are excluded from native latency, and normal allocator
metrics are unavailable.

The validated [instruction analysis](../results/change-0535/instruction-analysis.json)
has SHA-256
`c6ee0403185874146be22570f245f0fb07f70d6e18cfbb80c12914672d5da4ea`,
schema `litchi-ole2-change-0535-instruction-analysis-v1`. It validates 22 raw
profile dumps (20 timed and two setup), one relocation bias of `0x0`, 313
collector instructions, and exact equality between mapped instruction Ir and
collector function self Ir. Its direct-call and jump metadata remain separate
from instruction self cost.

## Native baseline

The native lane contains two warm synthetic in-memory cases, two repeats, 20
warmups per case, and 1,000 retained samples per case, for 4,000 samples.
Times below are nanoseconds. They are repeat-1/repeat-2 observations from the
same baseline binary, not before/after results.

| Case | Repeat 1 p50 / mean / p95 / p99 | Repeat 2 p50 / mean / p95 / p99 | p50 repeat drift |
| --- | ---: | ---: | ---: |
| `xls_owned_source_open_one_cell` | 95,615 / 96,685.790 / 107,160 / 112,910 | 95,590 / 96,794.535 / 107,590 / 114,000 | −0.0261% |
| `cfb_open` (`few-large`) | 68,840 / 69,357.530 / 74,410 / 76,200 | 75,790 / 76,262.307 / 81,170 / 82,450 | **+10.0959%** |

The XLS-owned row's maximum changes from 119,580 to 179,681 ns
(+50.2601%), while its p50, mean, p95 and p99 stay within 0.97% across the
two repeats. The CFB few-large maximum changes from 92,420 to 99,110 ns
(+7.2387%). Maximum, tail, RSS and system-time observations remain
descriptive same-build records; no sample or outlier is discarded.

The normal binary does not report allocator metrics, and this batch has no
allocator-instrumented lane. Consequently, it makes no allocation-call,
allocated-byte, reallocation, or incremental-peak claim. Native timing is
the only wall-time evidence in this batch.

## Same-build variation review

The [analysis record](../results/change-0535/analysis.json) retains all 13
same-build flags. They are repeat statistics, not matched candidate/control
flags, and they do not establish a cause:

| Case | Retained metrics and repeat-2 change |
| --- | --- |
| `xls_owned_source_open_one_cell` | `timing_stats.max` +50.2601%; `timing_stats.standard_deviation` +21.6449%; `rss.system_seconds` +5.8824%. |
| `cfb_open` (`few-large`) | `timing_stats.confidence_interval_95.lower` +9.9694%; `.upper` +9.9413%; `timing_stats.max` +7.2387%; `.mean` +9.9553%; `.min` +8.4829%; `.p50` +10.0959%; `.p95` +9.0848%; `.p99` +8.2021%; `rss.system_seconds` +7.6923%; `rss.user_seconds` −5.2632%. |

The records are retained for context in the raw analysis. They do not permit
sample trimming, stable-tail claims, or a before/after attribution. The
13-flag count is independent of any future candidate admission rule.

## Profile and instruction attribution

The profile lane has two repeats for each selected case, five timed Callgrind dumps
per profile, and two CFB setup dumps classified outside the timed collection:
20 timed constructor dumps and two setup dumps in total. The historical 0534 [collector
attribution review](../results/change-0535/collector-attribution-review.md)
keeps the selected constructor's inclusive cost separate from each exclusive
owner and from direct callee cost.

Across the five timed XLS-owned dumps in each repeat,
`SectorChainScratch::collect_exact` has 1,120,228 exclusive self Ir per dump,
or 5,601,140 Ir in the five-dump sum. Across the five CFB few-large dumps,
the collector has 1,114,389 self Ir per dump, or 5,571,945 Ir in the sum.
These are per-workload profile observations, not a mixed-corpus aggregate
and not proof that the whole owner is removable.

The current 0535 direct edges are per-dump values, so the small growth-cost
variation remains visible rather than being replaced by the historical 0534
attribution:

| Profile and dumps | Collector self Ir | Visible direct children |
| --- | ---: | ---: |
| XLS-owned repeat 1, parts 1–5 timed | 1,120,228 each | 43,096 each = 41,806 `memset` + 1,290 `finish_grow` |
| XLS-owned repeat 2, part 1 timed | 1,120,228 | 43,060 = 41,806 `memset` + 1,254 `finish_grow` |
| XLS-owned repeat 2, parts 2–5 timed | 1,120,228 each | 43,153 each = 41,806 `memset` + 1,347 `finish_grow` |
| CFB few-large repeat 1, parts 2–6 timed | 1,114,389 each | 21,638 each = 20,807 `memset` + 831 `finish_grow` |
| CFB few-large repeat 2, parts 2–6 timed | 1,114,389 each | 21,638 each = 20,807 `memset` + 831 `finish_grow` |

The two CFB setup dumps each have collector self Ir 1,114,389 and direct
children 21,671 (20,807 `memset` + 864 `finish_grow`); they are excluded from
the timed totals. The current [instruction report](../results/change-0535/instruction-analysis.json)
maps the contiguous valid-loop range `0x2f29480`–`0x2f2952e` to 99.9325%
of XLS timed collector self Ir and 99.9723% of CFB timed collector self Ir. The `memset` edges correspond
to retained visited-map clearing; `finish_grow` is a lower-level growth path
reached by fallible reservations. These are current per-dump graph values,
and their labels can include collection-off context, so they are not
operation-local allocation or dynamic call counts.

The validated current mapping shows the visited logical-length and word
checks, mask load, duplicate test, successful OR, sector-vector capacity
check/store, and next-marker handling in the collector's contiguous ordinary
loop. There is no positive `collect_exact` edge to
`CheckedBitSet::contains`, `CheckedBitSet::insert`, `Vec::push`, or reserve in
the selected raw graph. `CheckedBitSet::insert` has a separate binary-wide
out-of-line record for other callers; the absence of a collector edge does
not establish a timing cause beyond the mapped code shape. The [captured
assembly](../results/change-0535/baseline/assembly-index.json) and symbols
bind this current mapping; no speedup or causal latency explanation follows.

The collector's self cost includes its required entry reset, empty/start/table
checks, exact reservations, visited-map preparation, checked index and cycle
checks, insertion, ordered push, FAT/MiniFAT lookup, terminal-marker checks,
and failure reset. Any future candidate must preserve these operations and
their error/allocation order. The separate MiniFAT and regular-FAT scratch
instances and collect-then-claim sequence remain required.

## Quality reuse and limits

Because the runtime and harness are unchanged from final 0534, the frozen
plan binds reuse of that exact source's 14 quality gates and 4,382 executed
tests after custody verification. The retained prior artifacts are the
[quality summary](../results/change-0534/quality-summary.json), SHA-256
`816f28d3926d119972dbc45de1ac8960d93cd6bc370eb1333dca802dfe6fb2b7`,
[verification](../results/change-0534/verification.json), SHA-256
`106c12e5e76a555186915df053a9b708c27d5836a6c17b963a9d6a238cec5db7`, and
[sealed inventory](../results/change-0534/SHA256SUMS), SHA-256
`ec37ed38e7e974a7b77659b5791fe2241d7c83dc3edb188de0ff59e548d5a6b4`.
This reuse does not add 0535 Rust test executions or turn the diagnostic into
an optimization result. The [0535 postcleanup verifier](../results/change-0535/verification.json) passed
all components, including exact prior-quality custody. Five in-memory verifier
tamper probes passed. Both owned paths are absent and the final seal is checked.

The selected files are warm synthetic in-memory XLS/CFB inputs. They do not
cover cold or range providers, native Office producers, physical-I/O
variation, concurrency scaling, fuzzing, or broad CRUD behavior. Callgrind
owner shares and assembly shape are mechanism evidence only. OLE2 and OOXML
remain the active priority; ODF is deferred until that goal completes, and
iWork is excluded.

## ADR boundary

The 0535 ADR manifest retains the 30 accepted ADR/index bindings. This
diagnostic preserves the existing crate topology and private CFB ownership
(0001, 0002, 0024), snapshot and validation behavior (0003, 0006), measured
performance and resource accounting requirements (0005), source/binary
verification custody (0008), and physical package/directory ownership
(0010, 0011, 0026). It adds no public API, dependency, unsafe code, provider,
allocation policy, or validation-policy change.

## Disposition

0535 is a source-bound baseline and instruction-attribution record. The CFB
repeat-2 p50 drift and the retained 13 same-build flags require review but do
not support a causal slowdown or speedup claim. The validated instruction
analysis refines where the collector cost is emitted, but it cannot by itself
authorize a runtime change. Any future collector candidate needs fresh
native evidence, independent contract coverage, and explicit allocation and
source custody guards before adoption.
