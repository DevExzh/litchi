# 0521: borrow XLSX validation events and namespace state

`performance_claim: scoped source-backed XLSX edit/save improvement`

`claim_authorized: true`

Source-backed scalar-cell XML validation now inspects slice-borrowed events
and the reader's current namespace resolver. It no longer converts each event
to owned storage or clones the resolver for each event. Stack names and the
first bound dialect remain owned. All parsing, namespace checks, error order,
candidate validation and semantic readback remain intact. This removes
temporary ownership work from the owner identified in [0520](0520-xlsx-source-edit-phase-attribution.md);
it does not fuse validation passes or revive the rejected 0514/0516 designs.

The identical baseline/candidate harness adds the canonical System allocator
observer around the existing staged-set/commit interval. The normal executable
records explicit unavailable allocation samples; a separate allocator build
records measured samples. No API, dependency, unsafe code, resource policy,
format support or preservation contract changes.

## Native results

Two fresh children per case/shape retain 1,400 native samples across both
stages. Primary children retain 100 samples after 20 warmups; guards retain
30 after 10. CPU affinity is 2. Total time sums open, planning, staged
sets/commit and publication; publication includes dropping its returned
snapshot. Setup, remaining handle destruction, reopen and oracles are outside
this interval. Input is the instrumented in-memory source with a fresh
editor/cache each iteration, not a filesystem or range provider.

| Primary shape / repeat | Baseline p50 ms | Candidate p50 ms | Improvement |
| --- | ---: | ---: | ---: |
| Medium / 1 | 27.701 | 25.646 | 7.42% |
| Medium / 2 | 27.967 | 26.192 | 6.35% |
| Dense-sparse / 1 | 53.964 | 49.866 | 7.59% |
| Dense-sparse / 2 | 53.341 | 49.706 | 6.81% |

Primary commit p50 improves 7.65–8.50%, while planning improves 11.80–14.60%.
Planning also uses the changed validator. The one-edit, managed one-percent,
and vendor-extension guards improve whole-operation p50 by 6.55–9.80%.
Whole-operation p50, p95, p99 and mean improve in every matched primary and
guard child. Exact sample vectors and descriptive within-child bootstrap
intervals are in [comparison.json](../results/change-0521/comparison.json).
Two children per shape are limited process replication, not cross-host
confidence or a general tail-latency guarantee.

All 23 adverse phase/metric changes over 5% remain explicit in
[flag-review.json](../results/change-0521/flag-review.json): eight open,
seven publication and eight excluded-reopen metrics. Examples include the
second medium primary publication p50 increasing 7.93% and first managed
dense-sparse verification-reopen p50 increasing 12.95%. Their causes are
unproven. Fifty-two same-build repeat variations also remain in the comparison.
The admission decision accepts the narrow change because the required
whole-operation statistics, instruction counts and allocation counts improve
repeatedly, with unchanged semantic/output/resource identities. No standalone
publication, reopen or general tail improvement is claimed.

## Allocations and instructions

Four separate allocator children per stage retain five samples each, without
warmup. The region covers staged sets plus `MultiSourceEdit::commit`, excluding
publication, reopen and oracles. Counts and bytes below are identical across
all ten samples per shape per stage.

| Shape | Allocation calls before → after | Allocated bytes before → after | Incremental region peak |
| --- | ---: | ---: | ---: |
| Medium | 267,592 → 118,744 | 31,104,325 → 20,344,427 | 2,984,983 unchanged |
| Dense-sparse | 512,507 → 225,771 | 48,020,312 → 27,331,029 | 7,335,225 unchanged |

Allocation calls fall 55.62–55.95%; allocated bytes fall 34.59–43.08%.
Reallocations remain 10,798/19,564. Live-byte balances reconcile exactly;
incremental peak is region peak minus live bytes at entry. The observer is
operation-global, so these are commit-region results, not validator-only
allocation counts. Allocator-build timings are not native latency evidence.
Whole-child RSS, including setup and verification, has no adverse change over
5%; observed changes span -5.03% to +3.06%. No RSS reduction is claimed.

Eight separate normal-binary Callgrind children isolate one selected timed
commit each. All preceding lifecycle dumps are retained. The selected fourth
dump has one positive direct runner edge, and raw summary equals incoming
inclusive Ir equals self plus direct callees. This scope excludes staged sets.
Commit Ir falls 10.35–10.42%, and validator inclusive Ir falls 32.01–32.07%.
Every baseline profile has a positive validator-to-`Event::into_owned` edge;
every candidate profile lacks it. Other parser ownership calls remain.
These are guest-instruction diagnostics; inclusive child rows overlap their
parents and collection-off call metadata is not a timed event count.
All eight completed profiles retain Valgrind `brk segment overflow` notices;
allocator conclusions come exclusively from the separate canonical observer.

## Correctness, evidence and remaining work

Eleven differential tests compare exact errors and admission/refusal against
the frozen old loop. Coverage includes Transitional/Strict namespaces,
prefix/default aliases, delayed Empty/End scope pops, rebinding, foreign and
unbound namespaces, malformed attributes/references, DTD, UTF-8 and broad
truncation sweeps. Unchanged policy helpers are shared by both test paths.
An initial pre-capture run exposed an incorrect fixture error expectation;
the corrected 11-test baseline pass and original failure are both retained.

Native oracles retain deterministic output, unknown-member preservation,
no-op, clear/remove, foreign/stale refusal and semantic inverse restoration.
The vendor-extension shape additionally exercises partial-sink failure.
Medium/dense-sparse one-percent workloads touch all four sheets, so this does
not establish selective-subset access. Logical source/cache and managed budget
counters remain equal; unchanged budget charges are distinct from fewer heap
allocations. Corpus, output, source and budget identities match across stages
and instrumentation lanes after normalizing only planned iteration counts.

The retained quality receipts cover formatting, 1,274 all-feature XLSX
library/integration/doc tests, four new/existing harness metric tests,
all-feature workspace check,
XLSX/harness clippy, XLSX rustdoc, crate boundaries and strict claim checks.
The [source review](../results/change-0521/final-source-review.md) finds no
blocker; all 30 previously read ADR/index hashes remain unchanged. ADR 0003
source/patch semantics and ADR 0006 validation remain intact; ADR 0005 is
supported by source-bound native, allocation and instruction evidence.

The bundle verifies every raw receipt, serial interval, exact source-patch
replay, report and deterministic annotation, plus four negative semantic
vectors. See [reproduction and scope](../results/change-0521/README.md).
No fuzz, native Office-producer, physical FileSource/range, cold-cache,
hardware-counter, parallel-scaling or broad coverage-catalog completion is
claimed. Candidate reconstruction and worksheet layout/rewrite remain large
owners and require a new work-elimination proof before further changes.
Planning and publication remain alternative measured targets.

OLE2/OOXML optimization remains active. ODF is deferred until that goal is
complete; iWork remains excluded.
