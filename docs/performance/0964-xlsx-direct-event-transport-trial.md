# 0964 — XLSX direct event transport trial

**Decision: reject the screen.** The valid epoch-2 packet has twelve
descriptive native primary medians at or below the `0.95` signal threshold,
but these are screening signals without confidence intervals, not qualified
effects. Two native ingress-p99 medians exceed the `1.05` safety
threshold. The selector reports `reject_screen`, `adoption_allowed: false`,
and `statistical_effect_claim: false`. No production source change, full
confirmation, or adoption follows; production remains unchanged.

## Hypothesis and candidate

The retained 0962 frame-pointer diagnostic motivated replacing quick-xml
`NsReader` event transport with `Reader` plus a borrowed local
`NamespaceResolver` in exactly two XLSX loops: authoritative
`cell_values::validation::validate_xml` and raw worksheet
`parse_source_with_observer`. The hypothesis is reduced event-payload/tuple
transport; namespace binding, prefix resolution, validation, parser work,
and fallback authority remain.

The external frozen candidate is based on commit
`93d32552d98586f42822b2dc9bbec7c8ef6cafc9`. Its exact allowlist is:

* production: `crates/litchi-xlsx/src/cell_values/validation.rs` and
  `crates/litchi-xlsx/src/raw/worksheet/codec.rs`;
* test-only: `crates/litchi-xlsx/src/cell_values/shared_traversal_tests.rs`
  and `crates/litchi-xlsx/src/cell_values/validation_borrow_tests.rs`.

The candidate remains at
`/home/zhuhe/litchi-performance-0964/source-candidate`; the repository
production checkout is unchanged and public API, executor, cache, snapshot,
publication, and OPC ownership behavior are outside the candidate.

The transport review found the quick-xml 0.41 state machine preserved:
pending pops precede reads; `Start`/`Empty` declarations are pushed before
event checks; `End` stays in scope and `Empty`/`End` pops are deferred. Typed
namespace errors and priority, BOM offsets, caps, cancellation, borrowed
lifetimes, MCE ordering, parser/fallback boundaries, source freshness, and
cleanup remain in force.
The source review is a static **PASS**, not a performance claim.

## Quality and mechanism evidence

The final QA receipt is
`/home/zhuhe/litchi-performance-0964/receipts/quality-0964-final-v1.json`
(`status: pass`, source verified after quality). It reports 33 focused tests,
1,400 full library tests, 2,139 all-features package tests (1,407
library/unit and 732 integration), and 732 integration-only tests, all with
zero failures. All-target check, warning-denied Clippy, rustdoc, scoped
formatting, and crate-boundary audit pass. Four earlier compile/fixture
attempts remain retained and are excluded from the final pass.

The 0962 diagnostic selected 12,319 stage samples. Within its exact
source-admission subset (7,645 samples / 32,395,631,006 period weights),
exclusive physical self period was 6.923% for `NsReader::process_event` and
4.451% for `NamespaceResolver::resolve_event`. These sampled shares motivate
the hypothesis only and do not predict elapsed time or an effect. Read-only
codegen shows direct `process_event` and `resolve_event` edges changing from
`1 → 0` in each changed caller while unrelated caller edges remain; it does
not establish that resolver work disappeared.

## Epoch validity and screen result

Epoch 1 is wholly invalid and excluded. Its universal CPU 0–31 launch-affinity
witness raced `taskset` startup: all 36 native and 18 allocator children had
zero matching witnesses, so its reducer computed no ratios. Receipt:
`/home/zhuhe/litchi-performance-0964/receipts/epoch1-invalid-v1.json`.

Epoch 2 uses corrected pre-exec affinity controls and a fresh complete
54-child capture: 36 native and 18 allocator children, with no retries or
replacements. The route census is medium one-percent (93 updates),
dense-sparse one-percent (178), and medium one-edit (1). The independent raw
audit passes all 54 oracle/custody checks, reports no pooled samples, and
reproduces the selector decision.

| case | edit p50 | total p50 | ingress p99 safety | native RSS | allocator live peak |
| --- | ---: | ---: | ---: | ---: | ---: |
| medium-one-percent | 0.892037 | 0.933596 | **1.111553** | 0.994076 | 1.000000 |
| dense-sparse-one-percent | 0.886116 | 0.928109 | 1.024725 | 0.996686 | 1.000000 |
| medium-one-edit | 0.887339 | 0.931374 | **1.107516** | 0.992401 | 1.000000 |

Values are candidate-over-baseline medians of six paired native screening legs
and three paired allocator repeats. The screen has no confidence interval and these values are not statistical
effect estimates. Allocator allocation calls, allocated bytes, and
`region_peak_live_bytes` are exactly `1.000000` in all three cases; allocator
RSS is near 1.0. The two bold ingress-p99
medians trigger the hard safety rule. Future ingress-tail work requires a new
hypothesis and fresh gates; it cannot relax the existing safety threshold.
Ordinary and fallback/malformed-input confirmation controls remain required
for any future candidate.

The epoch-2 screen report is
`/home/zhuhe/litchi-performance-0964/analysis-epoch2/screen-v1.json`
(SHA-256 `6b2680a0b906b6b6908a06338794954135f46277db18360369f7e3cbedfe0742`);
the selector is `/home/zhuhe/litchi-performance-0964/analysis-epoch2/selection-v1.json`.
The independent raw audit is
`/home/zhuhe/litchi-performance-0964/review/independent-epoch2-raw-audit-v1.json`
(status `passed`), with verdict `screen rejected by the bound rule; no full
confirmation is indicated; no adoption or statistical effect claim`.

Cleanup passed and freed 35,701,850,112 disk bytes. The epoch-2 plan and
corrected runner are `protocol/plan-final-epoch2-v1.json` and
`tools/screen_0964_v2.py` in the external packet.

## ADR constraints

| ADRs | disposition |
| --- | --- |
| [0001](../adr/0001-priorities-and-api-layers.md), [0005](../adr/0005-io-memory-and-performance.md) | Correctness, preservation, finite budgets, and measured named gates precede speed; diagnostic shares do not become timing claims. |
| [0002](../adr/0002-crate-topology.md), [0024](../adr/0024-current-topology.md), [0030](../adr/0030-lazy-opc-part-decode.md) | XLSX ownership, dependency direction, and OPC part retention remain unchanged. |
| [0003](../adr/0003-snapshots-edits-and-patches.md), [0031](../adr/0031-execution-context-budgets.md), [0032](../adr/0032-snapshot-derived-value-memos.md) | No snapshot/memo/cache or executor change; one worker and explicit context budgets remain. |
| [0006](../adr/0006-validation-security-and-compatibility.md), [0008](../adr/0008-migration-and-verification.md) | Namespace/error/MCE/BOM/limit/fallback semantics and staged independent verification remain required; rejection leaves production unchanged. |

Primary source, QA, and prior transport precedent are recorded in [the source
review](results/change-0964/source-review.md), the external final QA/candidate
receipts, `/home/zhuhe/litchi-performance-0964/review/codegen-candidate-v1.json`,
and the 0810/0823 workflow records. Those precedents do not substitute
historical timings for this screen.
