# Next OLE2/OOXML work after the rejected 0524 CFB candidate

`scope: bounded read-only next-owner review using the fresh 0524 comparison`

`performance_claim: none`

Keep OLE2 and OOXML ahead of ODF until the OLE2/OOXML optimization goal is
complete. This review closes the 0524 visited-bit experiment and selects the
next measurement/design target. It authorizes no production change, and it
does not claim that the selected target is removable work.

## Decision

The next owner to measure is the source-backed XLSX cell-values one-percent
edit/commit route. Make `Snapshot::from_rewritten_source` the primary owner
and retain `raw::worksheet::edit::package::rewrite` as a separately
attributed companion phase. This is a larger OOXML opportunity than another
small CFB visited-bit or lookup tweak: the existing profiles identify two
large, independently named phases at the commit boundary, and the next
candidate can be retained only after a concrete redundant representation or
pass is demonstrated.

The intended route is the existing `MultiSourceEdit::commit` path in
[`source.rs`](../../../../crates/litchi-xlsx/src/cell_values/source.rs). It
rewrites the selected worksheet, builds a candidate snapshot, invalidates the
calculation chain when cell content changes, performs semantic readback for
every staged action, and publishes a reversible patch. The investigation must
keep those steps visible in the measurement boundary. It must not revive the
rejected source-pass or output-event fusion designs from 0514/0516, and it
must not turn a required validation into a presumed optimization.

## Why the 0524 CFB candidate is closed

The matched native gate rejects the candidate. The four primary XLS p50
deltas span `+1.46%` to `+4.87%` in repeat 1 and `-2.37%` to `+0.28%` in
repeat 2; no primary workflow reaches the required `3%` lower p50 in both
paired repeats. The CFB few-large guard is slower by `+2.35%` to `+2.41%`.
These are the frozen comparison's candidate-versus-baseline p50 deltas.

The fresh Callgrind comparison explains why the lower collector cost is not
enough to overturn that result:

| matched scope | baseline Ir | candidate Ir | change |
| --- | ---: | ---: | ---: |
| XLS owner, repeat 1 | 13,973,331 | 13,807,165 | -1.19% |
| XLS owner, repeat 2 | 13,971,462 | 13,808,665 | -1.17% |
| XLS `collect_exact` self, each repeat | 5,601,140 | 5,436,525 | -2.94% |
| CFB few-large owner, each repeat | 13,072,949 | 12,909,119 | -1.25% |
| all selected owners, aggregate | 82,640,159 | 81,944,616 | -0.84% |

Across the selected profiles, `collect_exact` self Ir falls from 23,885,540
to 23,187,490 (`-2.92%`), while its inclusive Ir falls from 24,665,846 to
23,967,966 (`-2.83%`). The selected owner still contains directory loading,
FAT/MiniFAT work, stream ownership, physical reconciliation, and the XLS
reader and parser. The candidate removes only the repeated checked-word and
mask calculation in the selected collector. It leaves `claim_sector`, the
required stream-allocation validation, physical reconciliation, and every
other `CheckedBitSet` user in place. A small diagnostic reduction therefore
does not imply a repeatable end-to-end latency, allocation, RSS, or I/O
reduction; the native gate confirms that it did not provide the required
workflow improvement.

The remaining CFB owners are not a safe automatic follow-up. Sector claims
establish ownership and error precedence; stream validation checks declared
chains and markers; physical reconciliation detects unclaimed non-free
sectors. The earlier chain review also rejected fusing `claim_chain` into the
collector because late cycle/marker failures and partial role mutation would
change observable behavior. A new CFB candidate needs a new proof of
removable work before it displaces the larger OOXML lead.

The supporting evidence is the [matched 0524 profile comparison](profile-comparison.json),
the [0524 profile scope review](profile-review.md), and the [0524 admission
matrix](README.md). Those reports mark Callgrind Ir as diagnostic only; the
native result remains authoritative.

## Why the XLSX reconstruction/layout boundary is next

The retained 0522 baseline profile names a substantial source-backed XLSX
commit owner on both current shapes. The two selected functions are sibling
direct callees of `MultiSourceEdit::commit`, so their inclusive edges are
disjoint within that selected call. They may be summed for this scope; the
other direct children account for the remaining commit work.

| shape / repeat | commit Ir | `Snapshot::from_rewritten_source` | worksheet `rewrite` |
| --- | ---: | ---: | ---: |
| medium / 1 | 197,179,569 | 121,438,255 | 75,197,769 |
| medium / 2 | 197,203,336 | 121,458,945 | 75,197,024 |
| dense-sparse / 1 | 377,116,324 | 229,721,741 | 146,606,245 |
| dense-sparse / 2 | 377,164,282 | 229,749,026 | 146,625,500 |

The reconstruction owner is about 60.91–61.60% of the selected commit Ir;
the worksheet layout/rewrite owner is about 38.13–38.88%. Together the two
sibling edges account for 99.72–99.79% of each selected commit profile. The
remaining direct children include action preparation, calculation invalidation,
readback comparisons, and destruction; those costs must remain separately
attributed.

The `5,601,140` to `5,436,525` `collect_exact` row above is the XLS
`xls-owned` profile total for each repeat's five timed dumps. It is not an
aggregate across the XLS and CFB workloads; the separate `23,885,540` to
`23,187,490` row is the aggregate selected-profile collector total.

The current implementation at
[`snapshot.rs#L686`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L686)
does the following in order: checks the source execution fence, validates the
rewritten worksheet XML, fully reparses it into semantic cells, clones the
source-backed snapshot state, installs the rewritten bytes as owned payload,
and checks the execution fence again. The package rewriter at
[`package.rs#L16`](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/package.rs#L16)
scans the original worksheet, validates the edit plan, computes layout effects,
reserves output, and copies or emits the preserved XML regions. A future
design may investigate whether one already-owned representation or pass at
this boundary is genuinely redundant, but the retained 0522 descendants give
the first concrete places to inspect. In its medium repeat-1 profile,
`from_rewritten_source` contains `raw::worksheet::parse` at 74,805,030 Ir and
`validation::validate_xml` at 46,624,585 Ir. The parser child contains
`Parser::parse` at 68,039,776 Ir, `semantic::materialize` at 5,780,824 Ir,
and `Parser::finish` at 5,668,345 Ir. The sibling `rewrite` edge contains
`scan_with_limit` at 71,007,022 Ir, including `Scanner::start_cell` at
21,000,853 Ir, `cell_address` at 13,164,862 Ir, `cell_tag` at 7,338,327 Ir,
and `Layout` destruction at 1,938,069 Ir. These are concrete descendant
owners to inspect in the retained [0522 medium annotation](../change-0522/baseline/profile-r1-medium.inclusive.txt),
not a claim that any of them is redundant. In particular, the full candidate
XML validation pass remains required unless a later design proves an
equivalent check with unchanged error order and resource accounting.

The accepted 0521 change already removed validator-to-owned-event conversion
work, and 0522's cell-reference candidate was rejected despite its diagnostic
instruction and allocation changes. The next probe must therefore target the
separate reconstruction/layout boundary. It must not repeat the cell-tag
scanner change under a new name or infer that full candidate validation,
reparse, readback, or calculation-chain handling is removable.

## Measurement prerequisite

First inspect the retained 0522 descendant reports and the current source for
the named children. The [0522 baseline profile analysis](../change-0522/profile-analysis.json)
and its [medium inclusive](../change-0522/baseline/profile-r1-medium.inclusive.txt),
[medium self](../change-0522/baseline/profile-r1-medium.self.txt), and
[dense-sparse inclusive](../change-0522/baseline/profile-r1-dense-sparse.inclusive.txt)
annotations already identify the parser, validation, layout scan,
cell-address, and layout-lifetime costs; read those descendants alongside the
implementations before scheduling new captures. The first design question is
whether a bounded changed-cell semantic update can reuse the existing snapshot
store after candidate XML validation, or whether a specific layout-scan
allocation/copy or parser subpass is provably duplicative. A proposed reuse
must account for formulas, shared formulas, rows, MCE/extensions, unknown
markup, limits, and all parser errors.

Only if that source-and-descendant inspection cannot distinguish a safe,
concrete removable work item should a new capture be scheduled. Such a capture
must be owner-scoped to the selected descendant or phase; do not queue another
whole `MultiSourceEdit::commit` baseline by default. If the audit identifies a
candidate, create a fresh before/after plan against the restored production
source using the existing medium and dense-sparse source-backed one-percent
edit cases, with the same source bytes, selected actions, limits, namespace
and markup-compatibility cases, output identity checks, and source/cache
oracles. Keep the operation clock around the existing staged-set/commit
interval, then add separate operation-local counters for:

1. worksheet `rewrite`, including its scan, plan validation, fallible output
   reservation, and emitted-byte construction;
2. `Snapshot::from_rewritten_source`, including candidate XML validation,
   semantic parse, state installation, and both execution checks;
3. calculation-chain invalidation, staged semantic readback, patch creation,
   and publication/drop, so a phase shift cannot be mistaken for removed work.

The allocator observer must cover the same commit region and report calls,
allocated bytes, live-byte balance, and incremental peak. The native lane must
report whole-route p50, p95, p99, and mean for two serially matched repeats per
shape, with setup, reopen, and oracles retained as explicit excluded phases.
Retain Callgrind or equivalent attribution only as a mechanism diagnostic;
also record copied/output bytes, logical source/cache counters, budget charges,
and whole-child RSS where the harness supports them. Use the existing
ABBA-style matched schedule and identity binding so a source or binary mismatch
cannot create a false phase delta.

Before a candidate patch is considered, the profile must identify a specific
allocation, parse, scan, or copy that can be eliminated or reused with an
independent proof. A lower child Ir value or a smaller intermediate vector is
insufficient. The candidate gate should require repeatable total-route p50
improvement on both primary shapes, no material allocation or incremental-peak
growth, unchanged output/source/cache/budget identities, and passing semantic,
error, preservation, and quality gates. Review every matched latency or RSS
change over the declared adverse threshold and every same-build variation over
5%; two children do not establish stable tails or cross-host confidence.

## Semantic and ownership guardrails

Any reconstruction/layout candidate must retain:

- exact source preservation for untouched worksheet and package bytes,
  including unknown attributes, namespace choices, extension and
  markup-compatibility content, and lexical details;
- candidate worksheet XML validation and a complete semantic reparse before
  publication, unless a later design proves an equivalent check with the same
  error ordering and resource accounting;
- all integer, cell, attribute, output, nesting, and aggregate-byte limits;
- source freshness and cancellation fences, typed allocation/limit/parse
  errors, and the established failure order;
- calculation-chain invalidation only for actual primary cell-content changes;
  semantic readback of every staged action; immutable source snapshots;
  reversible source-checked patches; and atomic publication;
- no public or dependency changes, unsafe code, hidden concurrency, broad
  cache, or format-normalizing fallback.

The investigation follows the applicable accepted contracts in the [ADR
index](../../../adr/README.md), especially ADR 0001's measured safety/API
layers, ADR 0003's immutable snapshot and atomic commit model, ADR 0005's
resource and measurement contract, ADR 0006's preservation and validation
rules, ADR 0008's evidence gates, ADR 0011's OPC ownership, ADR 0018's typed
calculation-chain ownership, and ADR 0024's current `litchi-xlsx`/OPC
topology. No ADR is amended by this review.

CFB `claim_sector` or physical-layout validation remains available only for a
future evidence-backed review that identifies required-check work which is
actually duplicative. DOCX/OPC publication is a later OOXML alternative if a
new current profile names removable work there; prior publication and source
reuse changes do not outrank this current XLSX owner. ODF stays deferred.

## Scope limits

This document is a queue decision and measurement prerequisite, not a
completion report. It makes no speedup, allocation, RSS, physical-I/O,
cold-cache, native Office-producer, provider, fuzz, scaling, or broad CRUD
claim. The current evidence uses the synthetic in-memory matrix and does not
complete those coverage requirements. No production file, test, build,
benchmark, or capture was changed for this review.
