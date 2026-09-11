# 0512: isolate current dense XLSX commit costs

`performance_claim: none; current operation-scoped instruction attribution`

`claim_authorized: false`

This batch establishes a current XLSX commit profile after the retained CFB
optimization. Production and harness Rust are unchanged at
`18d4d7fe1452bf43614a6efbd5feab4adb5a174f`. It replaces the old setup-inclusive
profile as the basis for selecting the next XLSX change; it is not an
optimization or a cross-era speedup claim.

## Scope correction and repeatable evidence

The generic commit/save runner constructs expected output with an untimed
commit and save. Its corpus generator also commits. Filtering only on the
commit symbol includes this setup, so it cannot establish timed-operation
cost. The fresh capture uses the commit-only case, zero warmups, collection
off initially, an end-anchored exact-owner commit toggle, and a reset at
`run_xlsx_update_commit` entry. Each tree shows exactly one runner entry and
three direct commit calls after reset. Open, edit staging, fixture generation,
final semantic readback, save and commit drop are outside collected costs.
Internal temporary destruction inside commit remains included.

The two profiles collect **11,907,741,960** and **11,909,172,574** instruction
references, a **0.012%** same-build difference. These are simulated CPU
instruction costs, not wall-clock phases or hardware cycles. Source and
artifact identities, full trees and raw profiles are retained.

| Disjoint direct commit child | R1 Ir | R1 share | R2 Ir | R2 share |
| --- | ---: | ---: | ---: | ---: |
| Source worksheet Store | 3,071,616,936 | 25.80% | 3,073,299,876 | 25.81% |
| Worksheet rewrite | 3,236,595,155 | 27.18% | 3,236,318,203 | 27.18% |
| Changed worksheet validation parse | 3,069,162,876 | 25.77% | 3,069,276,595 | 25.77% |
| Changed XML compaction | 2,426,402,056 | 20.38% | 2,426,410,798 | 20.37% |

The source and changed-input parse paths together dominate. Nested diagnostic
shares overlap the table: all `raw::worksheet::parse` costs are 51.55%, eager
Parser 46.64%, snapshot scanner 25.40%, cell address 5.13%, MCE preprocessing 4.73%,
and snapshot cell tag 2.71% in R1. Shared-formula resolution is only 0.05%.
The complete [attribution replay](../results/change-0512/attribution-summary.json)
retains both repeats. Do not add overlapping rows or interpret an inclusive
percentage as a removable fraction.

The reset leaves synthetic active-ancestor rows in the exclusive rendering;
those root rows are not summed or used for ranking. Descendant call metadata
can include calls made while cost collection is off, including final readback;
only the direct runner/commit and named direct commit-child edges establish
invocation counts. Both Valgrind logs report a `brk segment overflow` warning
and then exit successfully. This diagnostic allocator environment is another
reason to exclude instrumented timing/RSS and to avoid heap claims.

## Current native context

Two serial normal captures on CPU 2 use 30 samples after 3 warmups for each of
12 rows, retaining 720 durations. All native timer boundaries are unchanged:
commit-only times commit; commit/save times commit plus sequential write.
Expected-output generation, source opening, edit staging, sink reservation,
reopen/cell oracles and drops remain outside as applicable. The host is shared.
These 30-sample rows are descriptive, below the 500-sample registered latency
minimum, and provide no optimization confidence interval or speedup claim.

Median milliseconds:

| Case | Shape | R1 | R2 |
| --- | --- | ---: | ---: |
| `xlsx_one_cell_commit` | tiny | 0.163826 | 0.163311 |
| `xlsx_one_percent_commit` | tiny | 0.302166 | 0.302466 |
| `xlsx_one_cell_commit_save` | tiny | 0.206235 | 0.206125 |
| `xlsx_one_percent_commit_save` | tiny | 0.374187 | 0.372451 |
| `xlsx_one_cell_commit` | medium | 1.772457 | 1.758452 |
| `xlsx_one_percent_commit` | medium | 7.089053 | 7.030728 |
| `xlsx_one_cell_commit_save` | medium | 2.231949 | 2.245353 |
| `xlsx_one_percent_commit_save` | medium | 8.823105 | 8.902006 |
| `xlsx_one_cell_commit` | dense-wide | 108.735699 | 110.501998 |
| `xlsx_one_percent_commit` | dense-wide | 220.997433 | 223.119712 |
| `xlsx_one_cell_commit_save` | dense-wide | 155.170260 | 155.632553 |
| `xlsx_one_percent_commit_save` | dense-wide | 313.975840 | 314.185215 |

No same-build p50/mean/p95/p99 drift exceeds the frozen 5/5/10/15% thresholds.
Whole-child RSS is 139,236/140,776 KiB, including all setup, oracles and cases.
No exact document peak or memory improvement is established. Dense-wide is
two 256-by-256 sheets, 131,072 cells and 1,311 one-percent updates, with input SHA
`5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714`.
Dense one-percent output remains 388,095 bytes in 37 writes, largest 65,536 bytes;
one-cell output is 384,518 bytes in 35 writes. Catalog and sink identities match
preflight, both repeats and applicable diagnostics.

The whole-child hardware capture runs 30 dense commits after 3 warmups. Its
cycles/instructions/branches/branch-misses group has equal runtime and 100%
running coverage: 36,288,731,556 cycles, 141,401,785,247 instructions,
28,431,111,819 branches and 39,363,699 branch misses. It also records 59,045 page
faults, 102 context switches and 0 migrations. Setup, staging, generator/oracle
work and teardown remain inside this scope. These are not per-commit or
phase-local counts. Cache events are not collected; no cache or scaling claim.

## Decision and next implementation requirements

This evidence changes the next action: the two full parser paths and snapshot
layout are higher priority than the shared-formula loop or the remaining
plain-cell tag work. Removing an entire 25.40% scanner would bound simulated
commit improvement near 1.34x; real fusion can remove only part of that work.
This is a prioritization ceiling, not a latency prediction or a 10x claim.

Before attempting fusion, add allocator observations around the existing
commit and commit/save clock boundaries using the harness's current region
and `InProcessObservation` machinery. Normal binaries must report unavailable
allocation evidence, and only the separate allocator binary may supply those
vectors. A post-oracle boundary is also required for exact save-phase
attribution. Current generic save metrics measure sink activity only; missing
allocation, copied/decompressed/recompressed bytes are unavailable, not zero.

The implementation design must preserve semantic/style failure before snapshot
failure, original-byte offsets versus MCE-transformed input, x14ac fallback
precedence, exact no-ops, unknown/qualified attributes and lexical preservation.
Any temporary fusion state must have measured bounded memory overlap, and the
4,096-cell/1 MiB Store handoff stays unchanged. A smaller address/tag shortcut
remains a separately measured fallback, not a substitute for addressing the
larger bottleneck. 0472 already elides plain Tags; 0471's early buffer-release
experiment was rejected for lack of practical memory benefit and is not
silently revived. See the [source review](../results/change-0512/source-review.md)
and [scope review](../results/change-0512/scope-review.md).

## Reproduction and limits

Build: Rust 1.95.0 release, debug 0, incremental off, two build jobs; binary SHA
`213cb003e87ad35262ef9d71e313a836eb7b26b147d1060e94382b6e0b06c062`.
The source manifest is unchanged from 0511:
`4a0317050b0a28390a307b89321ad65d76d42a2a6d5619b6a5ba778db8378632`.
All 7,195 Rust/TOML/lock inventory files, two compile-time fixtures and all 30
previously read ADR files are bound. No production, dependency, API, unsafe,
owner, limit or format behavior changes. This evidence-only batch uses fresh
case oracles and metadata/tooling checks; it does not claim a new Rust test
suite, fuzz campaign or native Office validation run.

Formatting, crate boundaries, all ten strict registered claims, report
classification and Python syntax checks pass. The [evidence verifier](../results/change-0512/verify.py)
validates every raw duration, catalog/sink identity, source/build/fixture/tool
binding, exact profile edge, hardware group and a negative short-vector probe.
[summary.json](../results/change-0512/summary.json),
[gates.json](../results/change-0512/gates.json) and the
[final review](../results/change-0512/final-review.md) retain the results.
The [acceptance decision](../results/change-0512/acceptance.json) retains
attribution evidence only; no production optimization or new registered claim
is admitted. [verification.json](../results/change-0512/verification.json) and
[SHA256SUMS](../results/change-0512/SHA256SUMS) bind the final custody.

The [cleanup receipt](../results/change-0512/cleanup.json) removes only
`/tmp/litchi-goal-0512`: 1,768 files and 1,040,310,272
allocated bytes across unique file inodes, with no process references before
removal. The retained replay succeeds without executable or target artifacts.
Raw profiles, samples, catalogs, logs, source/build hashes and replay tools
remain under [change-0512](../results/change-0512/).

The broader performance goal remains open. OLE2/OOXML retain priority until
their full optimization goal is complete; ODF remains deferred and iWork
excluded. Default coverage (41 cases, 213 rows, 43 corpora), 18 mapped correctness-only
selectors and 10 registered claims are unchanged. Physical-cold, real-provider,
operation-memory and worker-scaling gaps remain open.
