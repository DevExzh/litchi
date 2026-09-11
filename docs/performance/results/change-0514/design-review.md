# Conditional worksheet parser and snapshot-layout fusion

Status: feasibility review against source `33a21e0f0`.

This review covers the proposed private fusion of the eager SpreadsheetML
semantic parser and the lossless snapshot `Layout` scanner for an ordinary
source worksheet. It does not approve a production implementation. The
candidate is feasible only if the fused pass remains a discardable transaction
local observer and the existing transaction order, source identity, and
fallback behavior are made explicit in code and differential tests.

The main unresolved admission conditions are the cost of speculative scanning
for a semantic no-op and the live-memory overlap between the semantic parser's
raw cells and the temporary layout. A successful changed-edit benchmark alone
does not establish either condition.

## Existing order that must remain observable

For a worksheet edit, the current path in
`crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs` has these
boundaries:

| Boundary | Current behavior | Fusion requirement |
| --- | --- | --- |
| MCE/x14ac preprocessing | `sheet.store()` calls `raw::worksheet::parse`; `raw::worksheet::parse` runs the x14ac preflight, then `process_ooxml`, then the semantic parser. | Run x14ac and MCE processing in the same order. A deferred snapshot error cannot replace either error. |
| Semantic worksheet parse | Parser event validation, shared-formula resolution, string-table loading, and raw-cell materialization complete before a `Store` is returned. | Run the semantic observer first for each shared event and do not expose a snapshot error until this whole boundary succeeds. |
| Style validation | `Worksheet::store` calls `validate_styles(&parsed)` before publishing a cold cache value. A hot cache returns without revalidation, which is existing behavior. | Preserve both paths. A deferred snapshot error loses to a cold style error and must not cause a hot cache to be revalidated. |
| Action projection | Hyperlinks, metadata, merges, defaults, cells, rows, and columns are projected against the `Store`. Unknown-cell payload edits can return `EditBlocked`; equal before/after states are omitted. | Complete the same projections before exposing a deferred snapshot error. |
| Effective no-op | When every projected action is absent, the transaction continues at line 1557 and never calls `edit::rewrite`. A requested same-value cell/row/column action still loads a cold `Store` before this check. | Drop the temporary layout and its deferred error. Do not publish a layout or alter source bytes. The extra cold and hot scanner work is an admission question, not a semantic free pass. |
| Ordinary rewrite | `edit::rewrite` scans first, then validates actions, computes extension names, and writes. | `rewrite_with_layout` must preserve this scan-before-`validate_actions` order. A layout scanner error is exposed at this same rewrite boundary only after the earlier rows above have succeeded. |
| Derived input | Merge removals rewrite the source first, so ordinary rewrite then receives derived `after` bytes. | A layout built over `before` is never reused for a derived buffer. Merge-container and MCE paths remain separate passes. |
| Post-write verification | Changed grid/merge output is compacted and parsed again, then styles and readback are checked. | Keep this parser/style/readback phase unchanged. The source layout is not a proof for the compacted output. |

The relevant current call sites are `transaction.rs:1359-1363`,
`transaction.rs:1459-1557`, `transaction.rs:1560-1589`,
`raw/worksheet/mod.rs:28-52`, and `raw/worksheet/edit/package.rs:16-25`.

In particular, the `scan` call in `rewrite` currently happens after
`sheet.store()` and action projection. Moving it into the source parse loop
must not move its error across those observable boundaries. The validation
performed inside `rewrite` remains after the scan; only transaction-level
projection is required to precede a deferred scanner error.

## Required source gate and ownership contract

The fused path should be an internal helper with a shape equivalent to
`parse_with_deferred_layout`, returning a semantic `Store` plus a private
transaction-local scan result. It must not change the public API, `Store`
ownership, or the persistent cache shape.

The safe gate is:

1. Run `x14ac::may_contain_descent`. If it is true, run the existing x14ac
   capture before processing; if it is false, retain the empty extension
   values and run that capture only in the existing plain-path parser-error
   fallback.
2. Run `process_ooxml` exactly as the current parser does.
3. Enter the fused reader only when the result is the original source bytes
   (`Cow::Borrowed`, or an equivalent pointer-identity proof). The proof must
   cover the pointer and length, not just equal contents.
4. Use the existing parser and snapshot scanner separately for an owned MCE
   result or any path whose byte positions do not refer to the transaction's
   source part.

`process_ooxml` can return borrowed bytes for an ordinary, MCE-free source and
owned bytes after MCE normalization. A byte search for an MCE namespace is not
an adequate reuse gate: processing can reject before parsing, and source
spans are valid only after the borrowed-result proof. `parse_defaults` has its
own preprocessing transaction and should not be folded into this candidate.

The private result needs a source token or equivalent guard. A practical
shape is a transaction-local wrapper containing the layout, the first scanner
error (if any), and an identity token for the exact `Arc<Vec<u8>>` used at the
rewrite boundary. `rewrite_with_layout` must reject or fall back to a fresh
`scan` if the source token does not match. It must never accept a layout made
over `before` when the input is merge-rewritten `after` bytes. Keeping an Arc
clone in this private wrapper is acceptable; putting a `Layout` in
`SheetData`, `OnceLock`, or another long-lived cache is outside this review's
scope and violates the bounded-cache intent.

The `Store` cache remains the only semantic cache. The layout and the source
scanner error are ephemeral values owned by one commit, and both are dropped
after a rewrite or an effective no-op. A failed scan cannot publish a partial
layout or a partial store. A failed semantic parse cannot be converted into a
successful scan result.

## Error precedence contract

The fused driver should implement the following precedence, with “scanner”
meaning any error that the existing `snapshot::scan` would return, including
layout finalization and event-limit errors.

* If x14ac capture is required by the existing preflight, its error remains
  first. For the plain path, if the semantic parse fails and the existing
  fallback capture then fails, that fallback x14ac error still replaces the
  parser error exactly as it does today. A saved scanner error is discarded in
  either case.
* `process_ooxml` errors remain before semantic parsing and before snapshot
  scanning. MCE-transformed input uses the existing separate path.
* A semantic parser error wins over a saved scanner error. This includes
  malformed XML/UTF-8, depth and context errors, cell/row/column validation,
  general-reference decoding, and leaf text checks. Call the semantic
  observer before the snapshot observer for an event.
* Errors after the event loop also win: shared-formula resolution, shared
  string loading, formula/value decoding, checked reservations, and every
  `materialize` error must complete before a scanner error can escape.
* `Worksheet::store` style validation retains precedence over a deferred scan
  error. The cached-store path must retain its current no-revalidation
  behavior.
* Action projection retains precedence over a deferred scan error. This
  includes merge conflicts and protected/unmodeled decisions made by the
  projection, defaults errors, unknown-cell payload blocking, and invalid or
  ineffective cell/row/column states. A requested same-value action is
  ineffective and must discard the deferred state without invoking rewrite.
* Once an effective ordinary plan reaches `rewrite_with_layout`, the
  deferred scanner error is returned at the same point at which `scan` would
  have returned it. If there is no scanner error, the existing
  scan-then-`validate_actions`/row/column/default validation order is retained.
* Compaction, post-write semantic parsing, style validation, and readback
  errors remain after the rewrite as they are today. They must not be hidden
  by the source scanner state.

This ordering is stronger than comparing successful output bytes. A useful
differential oracle compares the typed error variant and relevant context for
fixtures with two independently failing conditions.

## One reader, two observers

The semantic parser in `raw/worksheet/codec.rs` and the snapshot scanner in
`raw/worksheet/edit/codec/snapshot/scan.rs` both use `NsReader`, enable checked
end names, and resolve events through the reader's namespace resolver. A
shared driver is possible, but it must preserve the following properties:

* Read each event once, including comments, declarations, processing
  instructions, doctypes, CDATA, general references, and the final `Eof`.
  Capture the event start and end buffer positions at the same reader calls
  used by `snapshot::scan`; positions must be from the original byte slice.
* Call `resolve_event` once and pass the same resolved event, decoder, and
  resolver state to both observers. Resolving twice can advance or otherwise
  perturb namespace state. Do not reconstruct a resolved name from a local
  name or a decoded string.
* Preserve namespace scope and local rebinding for every event. Qualified
  element/attribute names, unknown names, and the snapshot's lossless tag
  bytes must all be derived from the event at its current resolver scope.
* Keep the current attribute behavior of each observer. The semantic cell
  scanner decodes unqualified `r` as it encounters it, ignores semantic
  prefixed attributes, and updates the inferred column before parsing style
  and metadata. The snapshot scanner then performs `cell_address` followed by
  `cell_tag`, normalizing all attributes and preserving qualified ones. A
  shared driver must not replace these distinct loops with a common
  pre-decoded attribute map.
* Run semantic callbacks before snapshot callbacks. If the semantic callback
  returns an error, stop as the current parser does; discard the temporary
  layout and run only the existing plain-path x14ac fallback if required. If
  the snapshot callback returns an error after semantic success, save the
  first scanner error, disable scanner mutation, and continue semantic parsing
  to its normal completion.
* A scanner error must make the layout unusable even if it was collected up to
  that event. The fused result can retain the error for the rewrite boundary,
  but it must not call `finish_layout` on a partial structure or use partial
  spans for output.

The observers intentionally do not have identical semantics. Examples that
make callback order observable include:

* `Parser::scan_cell_attributes` can ignore a prefixed or unknown attribute
  value that `snapshot::cell_tag` must decode and preserve. A source-only
  malformed entity in such an attribute can therefore be accepted by the
  semantic parser and rejected by the snapshot scanner.
* The semantic parser decodes general references in relevant text targets;
  the snapshot scanner records payload spans and does not decode ordinary
  cell references. The semantic error must win.
* Snapshot `scan_guard` parses sheet protection and data-validation metadata
  that the semantic parser may leave as `Context::Other`. Its errors are
  deferred scanner errors, not new semantic errors.
* Both observers maintain depth, but a semantic leaf/context error at a
  nested event must win over a snapshot depth or tag error at that event.

The driver must not impose snapshot-only policy on the semantic pass. In
particular, the semantic parser currently declares `MAX_XML_EVENTS` but does
not count events, while `snapshot::scan` counts every loop iteration,
including `Eof`, and rejects before reading the event that exceeds its limit.
When the fused scanner reaches that limit, save the scanner event-limit error
and stop scanner work, but continue reading for the semantic parser. A valid
semantic parse over more than the snapshot limit therefore succeeds for a
semantic no-op and exposes the limit only for an effective rewrite. A semantic
parser error before or after the limit still wins.

The existing x14ac capture reader is a separate preflight. Its raw/active MCE
observer ordering and recovery rules must not be folded into this ordinary
source reader. Any source with effective MCE transformation, AlternateContent
selection, or source-offset loss remains on the existing separate passes.

## Cache, race, and source-version rules

`Worksheet::store` uses `SheetData.cells: OnceLock<Store>`. It returns a
prepopulated value without parsing and can perform duplicate cold parses when
independent readers race; the first successful `set` is the value published.
The fusion must respect that contract:

* If `cells.get()` is already `Some`, do not start a speculative semantic plus
  snapshot pass. Project actions from the cached store first; an effective
  ordinary edit can use the existing separate `scan`, and an ineffective edit
  must stay scan-free.
* On a cold path, publish only a fully parsed and style-validated `Store`, in
  the same `OnceLock::set` pattern. If another reader wins the race, use the
  canonical `get()` result as the transaction's store and drop the losing
  local store. Never publish or share its layout through the cache.
* The layout source token is independent of which racing store wins. It is
  valid only for the immutable source bytes held by this transaction, and it
  must be dropped if source identity or source version checks fail.
* Source-backed workbook paths retain their execution/source-version checks
  before and after parsing. A source change wins over any deferred scanner
  error. In-memory immutable packages do not need a new race protocol, but the
  private helper must not assume that a cache miss is exclusive.
* Do not turn the layout into a second weighted cache. ADR 0005's lazy cache
  and ADR 0025's bounded validated-store handoff do not authorize an
  unrestricted layout handoff. A successful changed commit may retain a
  validated post-write `Store` only under the existing 4096-cell/1 MiB limits;
  the source `Layout` is released before publication.

## No-op guards and required evidence

There are two different no-op cases, and both are required.

**Cold semantic no-op.** A source `Store` is uncached, and a requested grid
action has the same semantic state as the source. Current code parses and
style-validates the `Store`, projects the action, then skips rewrite. A
concurrent fused reader would nevertheless have paid for event dispatch,
layout spans/tags, scanner allocations, and possibly `finish_layout`. The
implementation must either avoid that speculative cost or measure it as a
separate behavior with an explicit regression gate. Dropping the layout after
projection proves only error isolation, not performance.

**Prepopulated-store no-op.** A read first populates `SheetData.cells`, then a
same-value grid action is committed. The action projection can be done from
the cached store before any candidate scan. This path should not launch a
fused reader and should retain exact no-op behavior even if the source has a
snapshot-only malformed construct.

The minimum guard matrix is:

| Case | Source condition | Action | Required result |
| --- | --- | --- | --- |
| Cold no-op | Semantic parse succeeds; snapshot tag/guard would fail | Same-value cell, row, and column actions | Successful exact no-op; no deferred error exposed; temporary state dropped. Measure the scanner work and peak overlap. |
| Hot no-op | Store prepopulated with the same source | Same actions | Successful exact no-op; no source scan or layout allocation. |
| Cold effective edit | Semantic parse succeeds; snapshot scanner fails | One effective ordinary grid action | Scanner error at rewrite boundary, after semantic/style/projection checks. |
| Semantic precedence | Materialization/shared formula/value failure plus scanner failure | Effective ordinary edit | Semantic/materialization error. |
| Style precedence | Cold style validation failure plus scanner failure | Effective ordinary edit | Style error. |
| x14ac precedence | Required x14ac capture failure plus scanner failure | Any edit reaching source parse | Existing x14ac error. |
| Projection precedence | Unknown-cell/default/merge/action projection block plus scanner failure | Invalid grid edit | Projection error before scanner error. |
| Metadata only | Source has scanner-only malformed grid content | Page/margin/setup/web/print/hyperlink action | No grid fusion; preserve metadata codec behavior and existing source-store validation. |
| Empty edit | No requested effects | Empty edit | Existing exact source return path. |

The cold fixture must contain a scanner-only failure that the semantic parser
accepts, rather than relying only on a valid changed worksheet. An unknown
qualified attribute whose value the snapshot tag decoder rejects is a useful
shape, provided the fixture is verified against the current parser and scanner
in the repository. The hot case must use the same bytes and prewarm the actual
`OnceLock`, so it detects accidental reparse/fusion rather than merely using a
second logically equivalent worksheet.

For successful no-ops, compare the change set and the source-part `Arc`
identity/bytes, not just the resulting cell value. For errors, compare typed
variants and context. Add allocator counters for cold and hot no-op cases;
the existing empty-edit test does not cover the requested same-value action.

## Memory and performance admission

The current two-pass cold path drops parser temporaries before the snapshot
scan. A fused path holds the semantic parser's rows/raw cells and the
snapshot layout at the same time. Layout tags, spans, formulas, merges, and
unknown-content records can make this overlap larger than either pass alone.

Use the change-0513 allocator fields as operation measurements, with the
incremental quantity defined as:

```
incremental peak = region_peak_live_bytes - live_bytes_before
```

Keep absolute peak fields separate from the incremental value. Compare the
control and candidate on at least cold effective edits, cold semantic no-ops,
prepopulated semantic no-ops, and representative large worksheets. Report
allocation count/bytes and peak live bytes by operation; a changed-edit
allocation reduction cannot pay for an unmeasured no-op regression or a
larger peak. Do not claim a process-wide heap reduction from operation-local
allocator samples.

If the fused layout/raw-cell overlap is materially larger, or if the
no-op guard regresses beyond the agreed performance budget, retain the
separate measured snapshot pass (or narrow the fusion gate) rather than
publishing an unconditional fused path. The fallback must still preserve the
same error-ordering and output differential checks.

## Differential and limit coverage before admission

The following cases should run against fused and unfused implementations and
compare successful bytes, semantic stores, and the error contract:

* Plain source with default namespaces, prefixed namespaces, strict and
  transitional names, local namespace rebinding, unknown qualified content,
  comments/CDATA/general references, rows/cells/defaults/columns, formulas and
  shared-formula groups, x14ac rows/defaults, and preservation tags.
* Snapshot-only malformed attributes or guard metadata combined separately
  with a valid edit, a same-value cold edit, a same-value hot edit, a metadata
  edit, and an action-projection failure.
* A source with more than `MAX_XML_EVENTS` that the semantic parser otherwise
  accepts. An effective grid edit must expose the snapshot limit at rewrite;
  a semantic no-op must not. Include a parser error in the same fixture to
  prove parser precedence.
* A nested unknown/leaf structure near `MAX_XML_DEPTH`, malformed names or
  attributes, invalid UTF-8/tag data, and namespace declarations that rebind
  prefixes. Verify source spans and preserved tags are byte-identical where
  output is successful.
* MCE/AlternateContent, merge-container payloads, merge removals/additions,
  and any source that causes `process_ooxml` to return owned bytes. These must
  prove the candidate is bypassed and the existing fallback remains the
  oracle.
* Concurrent cold `Store` reads that race `OnceLock::set`, plus a transaction
  using a cached store. Verify one canonical published store, no partial
  layout publication, and unchanged source ownership.

The event-limit test is especially important because adding a shared event
counter to the semantic parser would silently change a currently accepted
semantic/no-op input. The counter belongs to the deferred snapshot observer
only.

## ADR alignment and API classification

| ADR | Constraint applied here | Classification |
| --- | --- | --- |
| 0001, Priorities and API Layers | Preserve typed fallible errors, strict layer ownership, lossless unknown content, and no new public parser/layout surface. | Shared reader and `rewrite_with_layout` are private implementation details. |
| 0003, Snapshots, Edits, and Patches | Source snapshots stay immutable; commit validates before atomic publication; exact no-op behavior and source ownership remain observable. | Layout is transaction-local and discarded; it is not a snapshot cache or patch payload. |
| 0005, I/O, Memory, and Performance | Keep positional source identity/version checks, finite limits, lazy thread-safe cache behavior, and representative allocation/peak measurements. | `Store` remains the existing cache; no new persistent `Layout` cache is implied. |
| 0006, Validation, Security, and Compatibility | Preserve malformed/unknown content, typed blocks for unsupported semantics, checked limits, and MCE safety. | Snapshot failures are deferred validation state; they do not authorize bypassing semantic or MCE validation. |
| 0008, Migration and Verification | Keep XLSX as the vertical slice, remain continuously buildable, and require compile-tested negative/differential verification before advertising the optimization. | Candidate is an internal migration step with an explicit fallback, not a compatibility shim or new format claim. |
| 0011, OOXML Physical Package Ownership | OPC/package ownership remains separate from worksheet semantic and lossless rewrite ownership. | Source identity is taken from the owning immutable part; Layout does not cross package ownership boundaries. |
| 0018, XLSX Calculation-Chain Ownership | XLSX owns worksheet grammar; MCE processing and bounded validation remain in their existing owners. | Fusion does not move MCE, formula ownership, or post-write calculation-chain validation. |

The bounded validated-store handoff in
`docs/performance/changes/0025-xlsx-validated-store-handoff.md` is compatible
with this review only when its existing cell/byte bounds and source/style
identity checks remain in force. The x14ac ordering preservation in
`docs/performance/changes/0032-xlsx-no-extension-scan.md` is the closest
precedent: an optimization may defer work, but it must retain the old error
fallback when a later parser phase fails.

## Admission decision

Proceed with a private prototype only if it has the source-borrowed gate,
semantic-first/deferred scanner contract, source identity guard, and separate
MCE/merge fallback described above. Do not call the optimization complete
until the cold same-value and prepopulated-store no-op guards pass, the
snapshot event-limit behavior is differential-tested, and memory overlap is
measured against the control. If those measurements fail, the measured
separate scanner path remains the acceptable implementation for this slice.

No source, build, or benchmark changes were made during this review.
