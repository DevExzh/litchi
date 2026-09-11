# Implementation review: conditional worksheet parser/layout fusion

This source review describes the retained [candidate.patch](candidate.patch)
against base `33a21e0f0`, before restoration. No build, benchmark or test was
run by the reviewer. Root's final owner unit suite passes 966/966 tests,
including 17 new tests, and all-features library Clippy passes with warnings
denied. Source locations below refer to the archived candidate.

The principal error-ordering contract is structurally preserved. Subsequent
[admission evidence](admission-review.md) rejects the prototype: the cold
same-value path performs an unnecessary scan and holds the layout beside
semantic state. The structural discussion below explains the mechanism; it
must not be read as performance approval.

## What the patch actually does

`Worksheet::store_with_layout` is selected when the requested action contains
at least one cell, row, or column map and none of the metadata, merge,
relationship, or drawing maps. On a cache hit it returns the existing `Store`
and `None`, so a hot no-op and a hot effective edit keep the old separate
snapshot scan behavior. On a cache miss it obtains an Arc of the worksheet
part, calls `raw::worksheet::parse_with_layout`, style-validates the resulting
store, publishes only the store into `SheetData.cells`, and returns a local
`DeferredLayout`.

The fused raw path performs the existing x14ac preflight, calls
`process_ooxml`, and fuses only when the returned `Cow` is borrowed from the
same source pointer and length. The owned MCE path runs the ordinary semantic
parser and returns a deferred wrapper with no layout, so the later edit falls
back to the ordinary scanner over source bytes.

`Parser::parse_with_layout_with_event_limit_inner` reads one `NsReader` event
stream. It resolves the event once, calls `Parser::consume_event` first, and
then calls `EventScanner::consume`. A scanner failure drops the scanner state,
retains its first error, and lets the semantic parser continue. Semantic store
finalization runs before `EventScanner::finish`, so materialization and similar
late semantic failures win over layout-finalization failures.

At the transaction rewrite point, `DeferredLayout::rewrite` discards its
layout and deferred error for an empty plan, falls back to a fresh scan when
the source pointer/length does not match, returns a saved scanner error before
rewrite validation, or delegates to `rewrite_with_layout`. Merge-derived input
continues to use the ordinary scanner.

## Error-ordering result

The following parts of the implementation match the required contract:

* x14ac capture still runs before MCE processing when
  `may_contain_descent` is true. With no marker, the plain parser-error
  fallback still calls x14ac capture and can replace the parser error as in
  `raw/worksheet/mod.rs:48-54`. A scanner error is not present on either
  failure path.
* `process_ooxml` errors and UTF-8 errors occur before the fused parser can
  return a deferred layout. Owned MCE output never receives source-byte
  spans.
* Semantic event handling is visibly first in
  `codec.rs:390-398`. A parser context, depth, leaf text, general-reference,
  or attribute error returns before the snapshot observer sees that event.
* `finish_store` at `codec.rs:416` precedes `EventScanner::finish`. Shared
  formula resolution, shared-string loading, value decoding, checked cell
  reservation, and materialization therefore retain precedence.
* `store_with_layout` style validation at `workbook/model.rs:1485` occurs
  after the fused parse but before cache publication. A deferred scanner error
  is dropped if this validation fails. A hot cache retains the existing
  no-revalidation behavior.
* Transaction projection at `transaction.rs:1385-1575` completes before a
  non-empty ordinary plan reaches `DeferredLayout::rewrite`. Unknown-cell
  blocking, defaults errors, merge projection errors, and ineffective actions
  can therefore win or discard the saved scanner error.
* The `rewrite_with_layout` handoff retains the old snapshot order: the saved
  scanner error is returned before `validate_actions`, row/column validation,
  default validation, extension planning, and output allocation. This matches
  `rewrite`'s existing scan-before-validation order.
* Post-compaction parsing, style validation, and readback remain on the old
  path. The source layout is not reused for derived bytes.

The error-ordering conclusion assumes the current call graph and the source
identity check described below. I found no direct production violation of the
x14ac, semantic/materialization/style, or transaction-level projection order
in this patch.

## Reader, namespace, and limit result

The `EventScanner` extraction preserves the standalone snapshot loop's
checked-end-name configuration, event positions, event classes, depth check,
stack finish check, and `MAX_XML_EVENTS` boundary. In the fused loop,
`before_read` runs before the read exactly as in the old scanner. If its limit
fires, scanner state is discarded and the semantic parser continues reading
to its own normal result. This avoids adding the snapshot event limit to the
semantic parser, which currently does not enforce that limit. A valid semantic
no-op over the snapshot limit can consequently succeed; an effective rewrite
returns the saved snapshot limit error at the old rewrite boundary.

The shared dispatch also preserves the important namespace details:

* `resolve_event` is called once, and both observers receive the resolved
  namespace, decoder, and resolver state for that event.
* Event spans are captured around the same reader calls as the old scanner,
  and only the borrowed MCE-free source path can use them for output.
* Semantic cell attribute scanning remains first. It decodes unqualified `r`,
  performs coordinate and row checks, and can ignore unknown qualified values
  before the snapshot `cell_tag` decoder sees them. This keeps scanner-only
  malformed attribute errors deferred.
* Formula, merge, text, CDATA, general-reference, comments, declaration,
  processing-instruction, doctype, and EOF dispatch remain separate from
  semantic handling. In particular, a semantic general-reference error is not
  replaced by the scanner's payload marker.
* The scanner's guard parsing for protection and data validation stays a
  snapshot error and is not accidentally promoted to a semantic parser error.

The event-limit test seam now exists under `cfg(test)` and uses the same inner
driver with a caller-supplied limit. That seam is appropriate for exact-limit
and parser-after-limit cases. It should remain test-only; production must keep
the fixed snapshot limit and the semantic parser's existing limit behavior.

## Source identity and MCE result

`DeferredLayout` retains an Arc source and checks pointer plus length before
passing a layout to `rewrite_with_layout`. The transaction obtains the later
`before` Arc from the same immutable part, so the current in-memory call path
has a stable identity proof. A different allocation takes the ordinary scan
fallback, which is the safe behavior for the direct source-level seam.

The visibility hardening in the final source closes the earlier API-surface
concern: both the `rewrite_with_layout` function and the edit-facade reexport
are restricted to `pub(in crate::raw::worksheet)`. The function still accepts a
bare `Layout` and byte slice, so the exact source identity check remains in
`DeferredLayout`; the ordinary package wrapper freshly scans the same `content`
before calling it. Within that restricted call graph, an unrelated layout
cannot enter through the former wider crate-visible facade. Keep the wrapper
check in place if this seam is refactored further.

The borrowed check in `raw/worksheet/mod.rs:118-125` is the right conservative
MCE gate. It checks the actual `Cow` result rather than guessing from a byte
marker. MCE-transformed bytes and all offsets derived from them stay on the
separate parser/scanner path. x14ac capture remains outside the shared reader,
which preserves its raw/active MCE observer behavior.

## Cache and race result

The cache behavior is compatible with the existing `OnceLock<Store>` contract:

* A prepopulated store returns before any fusion work. Effective hot edits use
  the old scan, while hot no-ops never scan.
* A cold parse style-validates before `OnceLock::set`, as the old path does.
  If another reader wins the set race, `get()` supplies the canonical store.
  The local deferred layout is not inserted into the cache and remains tied to
  the immutable source Arc, so it is safe to drop after the transaction.
* No `Layout` field was added to `SheetData`, `OnceLock`, or a workbook
  snapshot. The existing bounded validated-store handoff remains separate.
* This patch changes the in-memory worksheet model only. Any source-backed
  execution/source-version checks must remain around their own parsing path if
  fusion is ever extended there.

The race test exercises concurrent cold reads and a commit and checks the
source and committed stores. It does not prove allocation behavior for a
racing fused parse, but it covers the important no-partial-layout and
canonical-store ownership rule at the API level.

## Admission condition: cold semantic no-op work

The transaction's eligibility check at `transaction.rs:1366-1377` uses the
requested maps, before semantic effectiveness is known. Consequently this
sequence still occurs for a cold same-value cell, row, or column action:

1. `store_with_layout` parses the whole source with the scanner and builds or
   finishes a `Layout`.
2. Style validation and store publication complete.
3. Projection discovers equal before/after state and omits the action at
   `transaction.rs:1517-1519`, `1535-1537`, or `1551-1553`.
4. The all-effective-actions check at line 1575 skips rewrite and drops the
   temporary layout.

The result is semantically correct and the scanner error is not exposed, but
the scanner traversal, tag/span allocations, layout finalization, and
parser/Layout live overlap have already happened. The hot same-value path is
guarded because `store_with_layout` returns `(store, None)` on a cache hit;
the cold case is still an unconditional speculative cost. Invalid projected
grid actions can pay the same cost before being rejected.

This is an admission condition, not a measured performance blocker and not a
reason to change no-op error semantics. The focused tests now cover the cold
semantic no-op and exact error parity, but tests do not observe scanner
allocations or establish a no-op budget. The performance harness should use a
source that the semantic parser accepts and the snapshot scanner rejects, so it
continues to prove error discard, and report that cold case separately from
changed edits. If measurement shows a material regression, the fusion gate
must be narrowed or the design must accept a separate semantic prepass; it
cannot be called free because the layout is dropped.

## Admission condition: live-memory overlap and evidence

The old cold path drops parser raw-cell temporaries before `rewrite` scans and
allocates its layout. The fused path keeps the `Parser` state and `Layout`
state alive together until semantic parsing finishes, so the candidate could
increase incremental peak live bytes even when it reduces traversal time.
The focused tests do not measure that overlap. Subsequent paired allocator
captures demonstrate the increase for cold same-value edits; see the separate
admission review for absolute and incremental values.

Admission requires aligned control/candidate samples for cold effective edits,
cold same-value no-ops, prepopulated same-value no-ops, and representative
large worksheets. Interpret the operation peak as

```
region_peak_live_bytes - live_bytes_before
```

and report absolute region values separately. Retain the separate measured
scanner path if the overlap or no-op cost is materially adverse. The 4096-cell
/ 1 MiB validated-store bounds do not bound this temporary source layout and
cannot substitute for a peak-overlap measurement.

## Test/evidence gaps to close

The current focused tests cover useful correctness slices: ordinary output and
unknown attributes, namespace rebinding and shared formulas, source-allocation
fallback, deferred scanner errors, semantic/style/x14ac precedence, MCE
decline, exact event boundaries, cold/hot no-op output, a hot effective edit,
and a concurrent cache read. The XLSX owner unit suite is now reported green at
966/966 with zero ignored; the x14ac `NaN` fixture is checked as the typed
`Error::Descent` variant against eager exact parity. Before performance
admission, close these evidence gaps:

* Add an allocator/harness guard for a **cold** same-value action. The existing
  empty-edit test does not enter `store_with_layout`, and a hot same-value test
  does not exercise speculative layout allocation.
* Assert no-op source ownership directly where the private test module can do
  so (`Arc::ptr_eq` on the worksheet part), in addition to empty patch and
  equal serialized bytes.
* Differentially compare typed errors for combinations of scanner-only
  malformed attributes with semantic materialization/shared-formula errors,
  style catalog errors, and action projection blocks. String checks alone are
  weaker than the ordering contract.
* Keep exact event-limit tests for parser errors after the limit and for a
  successful semantic no-op. Verify that no fused path double-counts the
  shared event stream.
* Exercise namespace rebinding and qualified attributes on both a successful
  changed rewrite and an ineffective action. Compare the full untouched
  source spans and semantic Store fields with the unfused oracle.
* Retain explicit fallback coverage for MCE-owned output, merge-derived input,
  source identity mismatch, and scanner finalization errors.

## Review decision

The production patch can serve as a correctness prototype: its dispatch order,
MCE/x14ac handling, deferred-error behavior, event limit, source allocation
fallback, and cache publication are aligned with the requested invariants.
The XLSX owner unit suite is green, but subsequent cold no-op and allocator
measurements reject retention as recorded in admission-review.md. The final
visibility restriction and `DeferredLayout` identity check preserve the
prototype seam contract; production is restored to base.

No source, build, or benchmark edits were made by this review.
