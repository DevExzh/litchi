# 0679: XLSX selected-scanner publication design

Status: retained, design only. No production code, public API, benchmark, or
cargo result is claimed by this record. `performance_claim: none`.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

## Authority and decision

Queue row 15 in [0672](0672-xlsx-stored-cell-allocation.md) leaves one half of
the stored-cell allocation work open: the cold eligible scanner retains one
record for every selected physical cell until the worksheet reaches EOF.
[0642](0642-xlsx-visit-cells-streaming.md) deliberately kept that vector. A
scanner that called the visitor as it encountered a cell would let a later
worksheet refusal arrive after an earlier callback, contrary to the existing
`visit_cells` contract and the goal's refusal-before-result rule.

The decision for this record is:

1. Keep the existing `cells` and `visit_cells` routes unchanged. A generic
   one-pass emitter to the user callback behind `visit_cells` is rejected
   because it publishes a prefix before the scanner has reached the refusal
   boundary. `cells` may investigate building its unreturned final vector
   internally, as described below. An additive source-scoped partial stream is
   a separate API option, with explicit callback and error semantics, rather
   than a silent route change.
2. Treat a private summary/replay scanner as the only in-memory shape that can
   remove all N retained semantic values while preflighting the complete
   eligible worksheet. Compact records or direct final-vector conversion may
   reduce memory without replay, but they do not eliminate all N retained
   records. This is a design seam, not an implementation authorization.
3. Do not use that seam behind today's `visit_cells` API without deciding two
   contracts first: whether a callback may run while a source reader is active,
   and how a replay read is made independent of a second archive or transport
   failure. With the current callback lifetime and no automatic plaintext
   scratch policy, two-pass replay is not justified for the existing API.
4. If the owner later wants the memory reduction, the compatible landing shape
   is an explicit source-scoped replay API (or an explicit scratch-backed
   provider) backed by the private summary/replay modes below. Changing the
   callback lifetime of the existing method requires an explicit contract
   record; an additive source-scoped method can document a different lifetime
   or partial-result rule under the standing authorization. Automatic plaintext
   scratch would conflict with ADR 0005; caller-provided scratch is already an
   explicit capability. This document does not silently make either change.

The relevant accepted decisions are [ADR 0005](../adr/0005-io-memory-and-performance.md)
(immutable positional input, hierarchical work budgets, lazy payloads, no
automatic plaintext temporary files, and measured memory/performance) and
[ADR 0006](../adr/0006-validation-security-and-compatibility.md) (complete validation,
typed refusal, and no partial publication of malformed known payloads). The
broader accepted ADR set and precedence are recorded in
[the ADR README](../adr/README.md).

## Current source path

The behavior that the design must preserve is visible in the current source:

* [`SourceWorksheet::visit_cells`](../../crates/litchi-xlsx/src/workbook/source.rs#L777-L839)
  resolves a complete `Selection` before the first callback. Its documentation
  says that the selected scan reaches worksheet EOF, resolves shared-string and
  style dependencies, validates every retained record, and runs without an
  active source reader during a callback.
* [`stream_selection`](../../crates/litchi-xlsx/src/workbook/source.rs#L895-L1030)
  opens a `PartView::with_verified_decoded_reader`, calls
  `raw::selected_worksheet::scan_range`, resolves dependency readers after the
  worksheet reader is gone, runs source and execution fences, and only then
  returns the retained `Selection::Selected` records. An ineligible result is
  discarded and sent through the materialized store.
* [`SelectedRecord`](../../crates/litchi-xlsx/src/raw/worksheet/selected.rs#L85-L99)
  contains an address plus either a semantic `Cell` or a deferred shared-string
  index. [`retain_selected`](../../crates/litchi-xlsx/src/raw/worksheet/selected.rs#L1230-L1239)
  reserves and pushes one record for every selected physical `<c>` record.
* [`Scanner::finish`](../../crates/litchi-xlsx/src/raw/worksheet/selected.rs#L501-L553)
  publishes an eligible `SelectedCells` only after the eligible stream has
  completed its XML/MCE processing and has checked root, sheet data, ordering,
  merge count and synchronization state. Change 0658 stops only the monotone
  ineligible branch at its marking event; an eligible worksheet still reaches
  EOF before publication.
* [`PartView::with_verified_decoded_reader`](../../crates/litchi-opc/src/source_backed.rs#L3656-L3693)
  supplies a fixed-buffer reader without caching the decoded part. A successful
  callback may consume a prefix, but the owner still drains and verifies the
  complete member. The implementation charges the declared uncompressed part
  size to `Resource::Work` for each invocation
  ([source](../../crates/litchi-opc/src/source_backed.rs#L3750-L3790)).
* [`ActiveFlow::Stop`](../../crates/litchi-ooxml-common/src/mce/stream.rs#L774-L890)
  explicitly drops the unread stream checks after a stop. That is safe for the
  monotone ineligibility verdict because the materialized parser owns the
  remaining worksheet validation; it is not a way to stop an eligible scan
  before all refusal checks have run.

The existing tests pin the observable boundary. The 0642 later-row fixture
requests `A1:A2`, places a malformed cell in row 9, and requires the same error
as `cells` with zero callbacks. The 0365/0366/0367 fixtures require zero
callbacks for malformed tails, CRC failure, dependency failure, source change,
and cancellation before publication. The stored-route callback test separately
allows a callback to stop after an observed prefix and preserves the callback
error, subject to the final source and execution fences.

## What “refusal before result” means here

For a cold eligible selection, no callback may observe a value until the
following have all completed successfully:

* SpreadsheetML lexical and semantic validation, including formulas, scalar
  values, inline text, shared-string indexes, styles, ordering, merges and
  selected-record shape;
* MCE/x14ac processing, XML structure and all configured input, output, depth,
  event and allocation limits;
* complete verified worksheet-member draining, including archive size and CRC
  checks;
* dependency selection and validation for shared strings and styles; and
* source-version and execution-context fences.

This is stronger than “the selected cells happened to parse.” A malformed row
after the requested rectangle is still a refusal of the whole range, and a
verified reader or source-version failure is still a refusal before the first
callback. A callback error is different: once a caller has explicitly observed
an admitted value, the existing API may return that callback's prefix error,
while the final source and execution fences retain precedence.

## Candidate seam: summary then replay

A future private scanner can make the publication boundary explicit with two
modes. The names below are descriptive; they are not proposed public types.

```text
SelectedSummary {
    requested: Rect,
    selected_count: usize,
    dependencies: SelectedDependencies,
    requested_shared_string_indexes: Vec<u32>,
    single_cell_merge_proof: optional bounded merge result,
    source revision/fence observation,
}
```

**Pass A — summary/preflight.** The scanner uses the existing semantic event
path and all current limits, but `finish_cell` validates a selected formula,
inline value, number, empty value, or shared-string index and then discards the
owned semantic payload. It increments `selected_count`, records only the
selected shared-string indexes needed by the dependency reader, and retains the
bounded merge/order state needed by the current validator. It must still
validate every selected and unselected semantic condition that can make the
eligible path refuse. The scanner reaches XML EOF for an eligible worksheet,
then the verified reader drains and checks the archive member.

After the reader closes, the owner fences the source and execution context,
sorts and deduplicates the requested shared-string indexes, and resolves the
same shared-string and style dependencies as `stream_selection` does today.
The summary is not a publication of cells; it is a proof that the source has
passed the first scan and that the dependency reads are ready.

**Pass B — replay.** The scanner opens a second reader with the same limits and
replays the worksheet in source order. It checks the summary's requested range,
selected count, address order, dependency maxima, and selected shared-string
index sequence. For each selected record it constructs one semantic `Cell` and
hands it to a private sink. At EOF it checks the replay summary and the final
source/execution fences.

The replay must not silently trust a count. A count-only check would let a
parser bug reorder cells while preserving the same length. The proof needs at
least a checked address/order sequence and dependency/index agreement; a
bounded rolling digest can replace retaining all addresses if the digest is
domain-separated and collision policy is explicit. The replay also needs the
same defensive “exactly one of semantic cell and shared-string dependency”
check that 0642 moved before its first callback.

This seam eliminates the long-lived `Cell` values from Pass A. It does not
eliminate the shared-string payload table: selected shared-string lookups still
retain `Vec<(usize, Text)>` until replay finishes. It also does not make
`cells` constant-memory, because `cells` must return an owning result vector;
Pass B could fill that final vector directly, but the result remains O(N).

## Safe reductions that do not change publication order

The complete vector can remain in place while smaller, contract-neutral pieces
are priced. These are compatible with the current refusal boundary and do not
need an active-reader callback:

* Replace the two independently optional payload fields in `SelectedRecord`
  with an internal tagged representation such as `SelectedValue::Cell(Cell)`
  or `SelectedValue::SharedString(u32)`. The current scanner's two constructors
  already create exactly one arm; the current both/neither checks then become
  an enum invariant. This can remove padding and the duplicate option state,
  but it still retains O(N) selected values. A `size_of` and allocation probe
  must establish the actual saving before changing the representation.
* Keep shared-string indexes in their narrow `u32` form through sorting and
  deduplication. `stream_selection` currently expands each selected index into
  a `Vec<usize>` before dependency resolution, even though
  `SelectedRecord::shared_string_index` is a `u32`. A private dependency seam
  that accepts validated narrow indexes, or a checked conversion only at the
  reader boundary, can remove that transient width without changing text
  values or fallback behavior.
* For `cells` alone, a future scanner can investigate writing directly into its
  final `Vec<SourceCell>` after the worksheet scan has completed, with a
  separate compact side table for deferred shared-string records. This cannot
  serve `visit_cells`, and shared-string records must not receive a placeholder
  value before dependency validation. It is therefore a separate, narrower
  allocation change rather than a streaming publication design.

Each reduction needs the same sequence/refusal differential, exact allocation
failure tests, and a dense/shared-string allocation probe. None permits a
callback before the complete selected scan and dependency fences.

## Why the seam does not fit today’s visitor silently

There are only three ways to deliver Pass B values:

1. **Call the user callback during Pass B.** This removes the selected-record
   vector, but the callback runs while the verified worksheet reader is active.
   The current `visit_cells` documentation explicitly promises the opposite.
   A callback can also re-enter workbook/package code while a source-backed
   archive reader is active; the lower-level archive contract warns that
   `ReadAt` callbacks must not re-enter the archive. This requires an explicit
   callback-lifetime and reentrancy contract, not a private refactor.
2. **Close the reader before calling the user callback.** Pass B must then
   retain every replay value or every encoded cell until EOF, recreating the
   same O(N) memory problem. A bounded chunk queue only moves the refusal
   problem: a late Pass-B refusal after the first chunk callback is still a
   partial result unless Pass B is already a source-independent proof.
3. **Reopen the worksheet once per cell.** This closes the reader before each
   callback, but repeats ZIP traversal and decompression O(N) times. For large selections it can exceed the configured work budget and defeats
   the measured purpose of this route; tiny selections may still fit.

The strict variant therefore needs a second input whose complete bytes are
already verified before callbacks begin. That could be the package's decoded
part cache or an explicit memory/encrypted scratch provider, but caching the
whole uncompressed worksheet trades the selected-record vector for a full-part
buffer and automatic plaintext scratch is forbidden by ADR 0005. Without that
source-independent replay, a transient second-pass transport or archive error
can still occur after a callback even though Pass A succeeded. A source-version
fence detects mutation, but it does not turn a later remote read failure into a
preflighted error.

Consequently, two-pass replay is a valid design **only** for one of these
explicit contracts:

* a new source-scoped visitor whose documentation permits callbacks during a
  verified replay and states the reentrancy, callback-error and final-fence
  rules; or
* a caller-provided decoded/scratch replay whose complete bytes are verified
  before the first callback and whose scratch lifetime is explicit.

Neither contract is present in `visit_cells`, so this method must not be routed
through an active-reader or partial-result replay silently. An additive
source-scoped method with an explicit callback lifetime and partial-result rule
can be evaluated under the standing authorization; it would be a distinct API
contract, not a hidden change to `visit_cells`. Keeping the existing method
all-or-nothing and reader-free therefore means keeping its selected record
vector for now.

An additive **partial stream** could instead make the one-pass emitter itself
the contract: the callback receives each scalar value while the worksheet
reader is active, and a later XML, archive, transport, source, or execution
failure may report an admitted prefix. That is a coherent seam when a caller
explicitly opts into those semantics, but it does not satisfy the existing
refusal-before-result prerequisite. It also needs a dependency policy. The
current scanner discovers selected shared-string indexes while reading the
worksheet, then closes that reader before
`stream_shared_strings_dependencies` opens the shared-string part; nested
verified readers are not available as a safe default. A partial stream must
either require shared strings to be pre-resolved/warm, preload the needed table
through a separate explicit capability, or use the summary/replay path. The
numeric/inline-only subset does not establish a general XLSX visitor contract.

## Cost model

Let `D` be the worksheet member's declared uncompressed size, `N` the number of
selected physical records, `K` the selected shared-string references, and `U`
the unique shared-string payloads retained for lookup. Let `Q` be the
separately charged shared-string/style dependency-reader work. The `D` and
`2D` entries below count worksheet passes only; total verified-reader work is
`D + Q` versus `2D + Q` when the resolved dependencies are reused during replay.
Reopening dependencies in replay adds their work again.

| item | current eligible selection | summary/replay |
| --- | --- | --- |
| worksheet decode, MCE/XML work and archive verification | one pass, approximately `D` declared `Resource::Work` | two verified passes, approximately `2D` declared `Resource::Work` |
| selected semantic retention | one `SelectedRecord` per selected physical cell, O(N) | summary retains count/dependencies and O(K) indexes; replay retains one transient cell per callback |
| shared-string dependency | O(U) `Text` payloads when needed | still O(U); not removed by the replay |
| `cells` owning result | O(N), required by its return type | O(N), even if Pass B fills it directly |
| source/ZIP reads and decompression | one full verified member read | two full verified member reads unless a verified replay buffer is supplied |
| callback lifetime | reader closed before first callback | active reader for direct replay, or O(N) buffering after reader close |
| budget behavior | a profile sized for `D` can admit one scan | the same profile must admit `2D`; the second reader must refuse before callbacks when budget is insufficient |

The current source does not separately measure `size_of::<SelectedRecord>()`.
The 0642 dense probe measured a different, useful bound: removing the second
`SourceCell` vector saved exactly 5,242,880 bytes (80 bytes × 65,536 selected
cells), while the cold selected-route peak stayed 14,615,454 bytes because the
scanner's own retained records still set the peak. That result demonstrates the
scale of the remaining allocation but must not be reported as the size of a
`SelectedRecord`.

The retained 0642 cold selected-route totals are 2,234,372 allocation calls,
72,410,333 allocated bytes, and 14,615,454 peak live bytes before the vector
copy was removed; after the change they are 2,234,371 calls, 67,167,453 bytes,
and the same 14,615,454 peak. The extra retained-record validation pass costs
about 852,085 instructions on the 65,536-cell probe (approximately 13 per
record) and the cold whole-sheet instruction result is about +0.19% to +0.20%.
These are evidence for the current boundary, not a forecast for two-pass
replay.

The cost is currently unpriced on real producer-shaped eligible input. The
0658 census covers 180 `.xlsx` packages and 326 worksheet parts in `test-data`:
all 326 are ineligible, with 325 `UnsupportedStructure` verdicts and one
`RichInlineText`. The selected eligible route is exercised by generated probes,
not by those 326 real worksheet parts. Any implementation claim therefore
needs a fresh marker-free corpus with sparse, dense, formula, inline, and
shared-string sheets.

## Correctness and boundary tests required before implementation

The future implementation must add focused tests before any performance claim.
The tests below are the minimum matrix; each row must compare the current route,
the summary/replay route, and the materialized `cells` oracle where applicable.

### Publication and refusal order

* A valid dense and sparse eligible sheet produces the same address order,
  values, count, and digest through `cells`, existing `visit_cells`, and the
  replay visitor, both cold and after the store is warm.
* A malformed cell or row after the requested rectangle produces the exact
  `cells` error and zero callbacks. Repeat with malformed root closure,
  missing `sheetData`, duplicate/out-of-order cells, merge-count mismatch, and
  a malformed selected record shape.
* Place every `NotEligibleReason` before the first selected record, on the
  selected record, and after the last selected record. The result must still
  take the materialized route with no partial selected callback.
* Exercise raw XML, MCE, x14ac, invalid UTF-8/entity, depth, event, input,
  output, and allocation-limit failures before and after selected records. A
  failure found by Pass A must produce zero replay callbacks.

### Archive, source and execution fences

* Corrupt worksheet CRC, compressed size, and uncompressed size metadata. The
  first pass must fail before replay; a replay transport failure must also be
  demonstrated. If the latter cannot be made pre-callback, the implementation
  must use a verified replay buffer or reject the strict API design.
* Change the source revision after Pass A and before Pass B: zero callbacks and
  the same source-change error as a one-pass read. Change it from a callback and
  verify the existing final-fence precedence and documented prefix behavior.
* Cancel before Pass A, between passes, and before each callback. A callback
  failure at positions 0, 1, and N must preserve the exact admitted prefix and
  final cancellation/source precedence.
* Call back into the package from a replay callback under an instrumented
  source. The existing API must continue to prove no active-reader callback;
  an opt-in API must state and test its reentrancy rule explicitly.

### Replay proof and dependencies

* Compare selected count, source order, every selected address, and the
  dependency maxima/index sequence between Pass A and Pass B. Include zero
  selected cells, one selected cell, repeated shared-string indexes, indexes
  outside the requested rectangle, and a selected missing/explicit-empty cell.
* Cover direct numbers, booleans, errors, formulas and cached values, inline
  text, empty values, shared strings, styles, and single-cell merged coverage.
  Test rich inline text, formula semantics, style inheritance, and general
  references that must remain ineligible.
* Exercise missing/out-of-range shared-string entries and style dependencies.
  The route must take the same typed fallback or refusal as `cells`; no replay
  callback may see a guessed text value.
* Force the defensive both/neither `SelectedRecord` shapes in a unit-level
  scanner test. Neither shape may reach a callback or an owning result.

### Bounded work and memory

* Set a work budget that admits one declared worksheet pass but not two. The
  second pass must return a typed resource refusal before the first callback.
  Set a budget at exactly two declared passes and assert the observed work
  count, then repeat with dependency-reader work included.
* Assert per-stream input/output/depth/event limits on both passes and the
  source part-byte limit before any replay publication.
* Measure allocations, peak live bytes, retained bytes, decompressed bytes,
  source reads, and instructions for dense, sparse, long-inline, formula,
  shared-string-heavy, and late-refusal sheets. Record `N`, `D`, `K`, `U`, and
  the selected output count with each sample.
* Include one profile with a verified decoded replay buffer and one without it.
  The buffer's memory/scratch cost must be reported; it cannot be hidden under
  the scanner's record count.

## Reproducible evidence already retained

The existing packets provide a repeatable baseline without changing production
code:

* [0642 retained results](results/change-0642/README.md) include the generated
  `dense-wide-probe.xlsx` recipe (`2` sheets, `256 × 256` integer cells,
  384,231 bytes), the four-way differential over 397 worksheets of 182 `.xlsx`
  files, allocation probes, and callgrind commands. The generated probe is
  eligible for the selected scan; it is not a real producer fixture.
* [0658 retained results](results/change-0658/README.md) include the 326-part
  real-corpus verdict census and the source-backed differential over 180
  packages. The census establishes that the real corpus does not currently
  price a future eligible selected scan.
* The 0642 focused refusal test in
  [`source.rs`](../../crates/litchi-xlsx/src/workbook/source.rs#L4017-L4050)
  is the minimum zero-callback witness. The 0365–0367 tests around the same
  file provide CRC, source-change, cancellation, dependency and malformed-tail
  witnesses.

A future probe should reuse the 0642 generated workbook and add: a sparse sheet,
a sheet with long inline text, a shared-string-heavy sheet with repeated and
unique indexes, formulas with cached values, and a late-refusal tail. It should
run a counting positional source and `OpcOperationAccounting` around one-pass
and two-pass scans, then record the summary dimensions above. No result from
that probe exists in this record.

## Disposition

The remaining prerequisite is answered as a contract boundary. A one-pass
selected scanner emitter is incompatible with refusal-before-result for the
existing method. A summary/replay scanner is technically concrete, but it only
removes every retained selected value if callbacks may run during a verified
replay or if a complete verified replay buffer is supplied. Both choices carry
explicit cost and contract semantics, while the ordinary two-pass path doubles
worksheet work and does not remove shared-string memory. The current no-reader,
all-or-nothing `visit_cells` behavior therefore remains unchanged. A future
implementation can first measure compact-record/direct-vector reductions, then
choose either an additive source-scoped partial stream or an explicit
scratch-backed strict replay if the memory result justifies its cost.
