# 0523: CFB open allocation instrumentation and current OLE2 attribution

Direct `cfb_open` lacked operation-local allocation observations. The harness
now brackets its existing constructor timer with the canonical allocation
region and publishes elapsed-aligned observations. Normal builds explicitly
report allocation metrics as unavailable. Fixture generation, the file-size
oracle, reporting and object drop stay outside the timer. Production code,
providers, public APIs, dependencies and validation policies are unchanged.

This is a measured enabler and current baseline, with no before/after speedup
claim. The complete [evidence and reproduction instructions](../results/change-0523/README.md)
bind source, binaries, corpus identities, commands and raw reports. Historical
0511 numbers are leads, not matched controls for this batch.

## Native and allocation baseline

Two fresh children per group retain 24,000 native durations across nine fixed
XLS workflows and three CFB shapes. The separate allocation build retains 720
operation samples. Its elapsed times are excluded from native summaries.

| Workflow | Native p50, repeat 1 / 2 | Allocation calls | Reallocation calls | Allocated bytes | Incremental peak bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| CFB tiny | 2.270 / 2.260 µs | 32 | 1 | 4,456 | 3,384 |
| CFB many-small | 136.070 / 135.560 µs | 544 | 7 | 214,186 | 193,026 |
| CFB few-large | 102.470 / 101.920 µs | 29 | 1 | 213,571 | 205,129 |
| `xls_owned_source_open_one_cell` | 130.495 / 131.086 µs | 126 | 25 | 223,774 | 191,084 |

Allocation columns are constant across all retained samples in both repeats.
Incremental peak subtracts entry live bytes from the region peak. Absolute
live-byte values, deallocations and whole-child RSS remain separately available
in [analysis.json](../results/change-0523/analysis.json). Every allocation sample
reconciles allocated/deallocated bytes with the entry/exit live-byte balance.

All seven same-build timing variations above 5% are retained in
[variation-review.json](../results/change-0523/variation-review.json). Eager
XLS one-cell p50/p95/p99/mean rise 5.98–6.83% in the second child; owned-source
open p95/p99 rise 5.40–5.90%; tracked-source open p95 rises 5.45%. Two child
observations on a shared host cannot establish their causes or stable tails.
No sample is removed and no cross-run confidence claim follows.

## Scope and next optimization

Eight Callgrind children retain five measured constructor calls each. The CFB
fixture-generation open has its own separate dump and is excluded from timed
attribution. XLS profiles end at the source-backed constructor, before the
selected-cell query. Positive incoming caller edges and exact instruction
accounting establish these boundaries. Whole-child grouped hardware counters
are usable in both repeats at 100% running time; their roughly 1.67/1.68 IPC
includes setup, queries, oracles, drops and reporting.

The current [profile analysis](../results/change-0523/profile-analysis.json)
attributes 40.08–40.09% of XLS constructor instructions exclusively to
`collect_exact`. Sector claiming is 17.81–17.82%, physical reconciliation
14.25–14.26%, stream-allocation validation 14.16–14.17%, and FAT loading
1.79–1.80%. These exclusive shares are disjoint; inclusive parent/child costs
must not be added. Direct CFB few-large spends 42.62% exclusively in chain
collection, versus 2.12% for tiny and 5.45% for many-small. The latter shapes
therefore remain useful guards rather than evidence of a large shared gain.

The [chain review](../results/change-0523/chain-review.md) proposes a private
checked visited-bit lookup/set experiment in `SectorChainScratch::collect_exact`.
It must preserve bit bounds, cycle/marker error order, scratch reset and
allocation failures. Ownership claiming and physical-sector reconciliation
remain required. Adoption requires a fresh matched native comparison with
useful repeatable workflow improvement and allocation/correctness guardrails.

CFB tests pass with all features and without default features (305 each);
the focused harness test adds one passing execution, for 611 valid test
executions across these configurations. Workspace compilation, CFB/harness
Clippy and harness rustdoc also pass.

The new focused harness test covers tiny/few-large reopen oracles, elapsed
alignment, allocator unavailable status and source not-applicable status.
The initial preflight test command passed but its build overlapped the final
test assertions; that source-custody mismatch is retained, and a second frozen
source run passed before capture. All ten quality gates and post-cleanup
evidence replay pass: 8,584 source files, 32 serial intervals, all three
analysis reports, 80 annotations and four rejected semantic corruptions.
Both owned build paths and Python caches are absent.

Synthetic in-memory inputs do not complete physical-provider, cold/range,
native Office producer, fuzz, concurrency scaling or broad CRUD requirements.
No fuzz campaign ran: cargo-fuzz and a nightly toolchain were unavailable.
OLE2 and OOXML performance remain the active priority; ODF is deferred until
that optimization goal completes, and iWork is excluded.
