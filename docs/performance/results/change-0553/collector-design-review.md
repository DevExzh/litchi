# Change-0553 commit-local compact collector review

**Status: draft-07 source-only review passes; no source blocker was found in
the five reviewed areas.**
This review is scoped to the proposed commit-local compact source-cell
collector for the source-backed XLSX `MultiSourceEdit` path. It does not approve
a performance claim or an adoption decision. The production baseline for a
fresh experiment is commit `8aa0c5baf0616d16c79eba0c6c28dc1716338ad6` (the
completed 0552 rejection), and all 0552 latency, allocation, RSS, planning, cap,
and quality gates remain unchanged.

The governing design record is
[`deferred-proof-feasibility.md`](../change-0552/deferred-proof-feasibility.md).
The relevant source authorities are the ordinary `litchi-xlsx` worksheet
parser, the complete edit scanner/writer, the source-backed snapshot
provenance checks, and the accepted ADRs in `docs/adr/`. A commit-local
collector is acceptable only as an optional optimization over those authorities:
it may return a temporary layout or decline, but it may not become a new
semantic parser, preservation authority, public error, or retained snapshot
state.

## Candidate packet findings

The immutable draft-07 packet is bound to
`8aa0c5baf0616d16c79eba0c6c28dc1716338ad6`. Its binding records the compact
collector hash
`ce3cab7a04e2a5be6de932c2227ee1915e890ddcd1c77480056828e2f27e9e1f` and the
complete source manifest. Draft-07 is therefore the source-review target;
later mutable candidate files are outside this finding. Its production
collector and writer are unchanged from draft-06; the only source addition is
the `#[cfg(test)]` attempt/acceptance counter hook and its private test module.

The draft-07 `MultiSourceEdit::commit` source uses a block around collection and
`rewrite_value_only_with_compact_proof` (`source.rs` around 1183--1196). That
block ends before aggregate accounting,
`Snapshot::from_rewritten_value_source`, staged-value readback, invalidation,
and publication. The rewrite result borrows only the immutable source slice;
the moved `Option<CompactLayout>` and its row/cell scratch drop when the rewrite
function returns, and the local `compact` binding is out of scope at the block
end. The wrapper also executes `drop(proof)` before either returning the compact
rewrite or entering the complete fallback (`package.rs` around 191--198).
This is the required source/store lifetime boundary; retaining the lexical
block and explicit drop is necessary because `CompactLayout` carries raw
identity facts rather than a Rust lifetime parameter.

An earlier compact source snapshot omitted the non-empty-cell span append in
`finish_cell`, which made the final row-cell cardinality check reject every
ordinary `<c>...</c>` record. Draft-07 records that span (compact.rs around
641--666), and its empty-child helper marks `<v/>` and `<is/>` without pushing
an unmatched stack frame. The source cardinality path is therefore live for
non-empty, empty, and primary-empty cells.

Draft-07 has an explicit `Event::Comment(_) | Event::Decl(_) | Event::PI(_) =>
true` arm. The direct unit case proves that processing instructions before,
inside, and after `sheetData` do not create collector-specific semantic state
and are accepted by collection. That case asserts collector acceptance only;
it does not assert writer-byte equality, so it is not byte-preservation
evidence by itself.

Draft-04's independent `is_core` predicate was a source blocker because it
accepted transitional and strict children independently. Draft-07 replaces it
with a root `SpreadsheetDialect` slot and `same_dialect` checks on every
SpreadsheetML start, empty, and end element (compact.rs around 34--41,
257--273, 455--459, and 615--619). The mixed-dialect refusal and prefixed
transitional/strict acceptance cases are present in the direct collector tests.

The draft-07 worksheet-order state now matches the scanner's direct-child
guards. `dimension` requires no prior defaults, columns, or `sheetData`;
`sheetFormatPr` requires no prior columns or `sheetData`; `cols` may follow
defaults but not `sheetData`; duplicate and empty `cols` forms decline. The
accepted `dimension, sheetFormatPr, non-empty cols, sheetData` case and the
scanner-error refusal table are covered by the immutable package differential.

The draft-07 `decode_attributes` still admits namespace declarations, validates
every attribute name as UTF-8, and decodes and normalizes every value with
duplicate/syntax checks. It intentionally rejects other prefixed attributes
(`xml:space`, relationship attributes, and extension-owned attributes), which
is a safe false-negative because the wrapper falls back to the complete writer.
The direct malformed-unused-attribute case returns `None`; no unused value can
bypass the complete attribute/namespace check.

The fixed 2 MiB accounting is a logical collector-state bound. Draft-07 charges
the collector/layout/reader envelope and every row/cell/stack capacity through
checked, fallible growth; conversion of those vectors into boxed slices does
not create a second capacity. Reservation overflow and equal-byte
source/store-identity refusal cases are present in the immutable resource
child. Temporary decoded attribute `String`/`Cow` values are not retained in
the layout and remain a later measured allocation/RSS concern, rather than a
source-review admission claim.

Structural text keeps the same explicit boundary. Whitespace-only text outside
`v`/inline `t` may be copied; non-whitespace structural text is a proof refusal,
never a new collector error. Draft-07 routes CDATA through a separate helper
that refuses `sheetFormatPr` CDATA, matching the ordinary parser, while
character references still conservatively decline. The direct collector test
covers the leaf-CDATA refusal alongside malformed unused attributes.

## Draft-07 concrete source handoff

This is the bounded source finding for immutable draft-07 production. It does
not infer performance, process memory, or admission and does not replace the
root-owned correctness and measurement gates.

* **Source/store bijection and order — pass.** `start_cell` and `empty_cell`
  compare each resolved address with the next authoritative sorted Store entry
  before advancing. `finish` requires the consumed entry count and cell count
  to equal `entries.len()`, and the row cell sum to match as well. Explicit and
  inferred row/cell forms use the shared scanner parse helpers; rows and cells
  are strictly increasing, with checked inferred row/column increments. The
  immutable direct cases cover explicit, inferred, empty, and identity forms;
  package coverage includes discontiguous cells.
* **Namespace, unused attributes, and fallback — pass.** The root dialect slot
  and `same_dialect` checks cover every SpreadsheetML start, empty, and end;
  mixed dialects decline while prefixed transitional and strict forms are
  accepted. `decode_attributes` checks all accepted start/empty attributes for
  duplicate and syntax errors, UTF-8 names, and XML 1.0 decoded/normalized
  values, including unused values. Other prefixed attributes are an intentional
  safe false negative. General references, DTDs, unknown or foreign elements,
  and malformed unused values return `None`, leaving the complete writer's
  diagnostics and precedence authoritative. The PI test proves collection
  acceptance only, not writer-byte equality.
* **Resource growth and lifetime — pass for source correctness.** The static
  envelope and every row, cell, and stack growth use checked accounting and
  fallible reservation; boxed conversion transfers the existing allocations.
  The proof retains spans and raw source/store identities only. The collection
  block ends before candidate readback, and `drop(proof)` precedes compact
  return or complete fallback. The resource child covers reservation and
  identity. Its 4,000-row case compares direct compact and complete bytes,
  omission spans, reduced bytes, and parsed address/cell/style state; its
  12,000-row metadata refusal compares the wrapper fallback against the same
  complete result. Transient decoded attributes and output allocation remain
  measured concerns.
* **Writer, dimension, and actions — pass for the proven boundary.** The
  compact plan accepts existing scalar `Update` actions, including the target
  plain `SetFormula` representation (`Payload::Set(Content::Formula)`), and
  direct action cases cover discontiguous updates, clear/clear-if-present,
  style, and the dimension close. Source `<f>` cells, shared/array/data-table
  formulas, and shared-string provenance decline before compact rewriting.
  Compact and complete bytes, omission spans, and reduced/full readback are
  differentially checked for the covered actions.
* **No-op and planning — source pass.** The commit path expands staged values,
  checks `actions.is_empty()`, clones the snapshot, and only then calls the
  collector; planning and snapshot construction are untouched. The collector
  is absent from `Snapshot`, `SourceState`, `Patch`, and publication state.
  Draft-07's private `#[cfg(test)]` thread-local counter observes the real
  `MultiSourceEdit` seam: empty and effective-no-op commits record `(0, 0)`,
  while a changed commit records one attempt and one accepted proof. This
  closes the premeasurement no-op coverage item without changing release
  behavior.

**Source-only verdict:** Draft-07 has no remaining source implementation
blocker in the five reviewed areas. Root may continue to the existing
correctness, quality, and measurement gates. This verdict makes no performance
or admission claim; runtime outcomes remain pending root's receipt.

The compact proof entry point bypasses the complete scanner's
`validate_actions` when it accepts a proof. Therefore the collector must either
prove the equivalent edit guards or decline before returning a layout. In
particular, protected sheets, data-validation ranges (including extended
validation), covered merges, markup-compatibility cell payloads, and formula or
shared-formula dependencies must use the complete writer so the existing typed
block errors and precedence survive. A source/store address match alone is not
enough to authorize compact output.

The action wording needs one deliberate distinction. The public plain
`SetFormula` operation is staged as `Action::Update` with
`Payload::Set(Content::Formula)`, so it may use the compact writer when its
target is an otherwise eligible existing scalar cell and the direct walk proves
that source cell has no formula or uncertain payload. A source `<f>` cell,
array/data-table/shared-formula structure, or shared-string dependency must
decline to the complete path. Tests must exercise the value-to-plain-formula
case separately from source-formula and shared-formula fallback cases.

## Required placement and lifecycle

The collector must be created only in `MultiSourceEdit::commit`, after the
existing staged actions have been expanded and effective-action eligibility has
been established. A sheet with no effective action must clone its snapshot
without walking the source. Planning (`edit_sheets`), snapshot construction,
and no-op behavior must remain unchanged. The single-sheet `SourceEdit` path is
outside this experiment unless separately named and measured.

The result must be an ephemeral, borrowed layout held only until that sheet's
rewrite finishes. It must not be stored in `Snapshot`, `SourceState`, `Patch`, a
candidate snapshot, or any later publication object. The collector must not
materialize a second semantic `Store`, call the ordinary parser merely to
obtain offsets, or replay planning. Its source slice and the authoritative
`Store` entries must remain owned by the same immutable snapshot. Before any
span is read, the candidate must verify the source allocation identity and
length and the entries allocation identity and length; a mismatch must decline
to the complete writer.

The candidate must preserve the existing source eligibility boundary: the
8 MiB source ceiling, UTF-8 requirement, MCE/`AlternateContent`/`x14ac`/
`dyDescent` exclusions, and the 131,072 provisional-event ceiling. A direct
collector may decline earlier. Its own dynamic storage must use checked
arithmetic, fallible reservation, and an explicit logical cap; this cap must
be documented as collector metadata only and must not be presented as a
process-RSS or total workflow-memory bound.

## Direct walk and XML authority

The intended walk is a borrowed `NsReader` pass over the original worksheet.
For every event, offsets must be taken from the reader before and after the
event and checked for conversion, ordering, and bounds. The walk must accept
exactly one transitional or strict SpreadsheetML root and keep namespace
resolution active for every start, empty, and end element. Unbound, foreign, or
mixed SpreadsheetML dialects, mismatched closing names, duplicate roots, DTDs,
and incomplete roots must decline. Prefix spelling may be retained in source
spans, but it must not change the resolved namespace or the semantic address.

The collector must not use local-name matching as a substitute for namespace
resolution. It must not widen the compact route to elements that the existing
validator rejects, including foreign elements, markup-compatibility content,
extension-owned content, formulas, shared strings, rich inline payloads, or
unknown direct cell children. Those cases must take the complete writer path.
The source eligibility filter is only an optimization admission check; it is
not permission to skip structural checks that the complete scanner performs.

Every `Start` and `Empty` element that can reach the complete scanner must have
the same XML attribute safety treatment. In particular, the collector must
iterate attributes with duplicate/syntax checks, validate UTF-8 attribute names,
and run XML 1.0 `decoded_and_normalized_value` on every attribute, including
attributes whose values are not needed by the collector. This is necessary for
the late malformed-unused-attribute cases: planning validation may accept an
allowed attribute that the complete writer later decodes. A collector decode,
namespace, structure, or equivalence uncertainty must return `None`, after
which the complete scanner/writer returns the established diagnostic and
precedence. The collector must never surface a new error, silently ignore a
latent malformed value, or publish compact output after skipping the complete
check.

The collector may retain no decoded tag or unchanged attribute metadata. It
must only retain the source spans and the minimum envelope data required by the
existing compact writer: the direct `sheetData` envelope, each row envelope,
each cell envelope, primary-child spans when the writer needs them, and the
optional preceding `dimension` span/reference. Any tag materialization for a
changed cell must remain bounded and follow the existing writer's lexical and
namespace behavior. Unchanged bytes, comments, whitespace, attribute order,
prefixes, CDATA, and omission provenance remain source-backed.

## Address, ordering, and cardinality proof

The source order must be tied to the authoritative store by a bijection. Let
`S` be the source cell sequence in worksheet order and `E` be
`snapshot.cells().entries()`. Since `Store::from_unsorted` exposes entries in
address order and does not retain source offsets, a count or an address lookup
alone is insufficient. For every cell event `i`, the collector must require:

```text
resolved_address(S[i]) == E[i].address
```

It must reject an omitted source cell, an extra source cell, a reordered source
cell, and a same-count address mismatch. At `sheetData` close it must require
that the consumed entry count equals `E.len()`, with no pending row or cell.
The source and store lengths must be checked before indexing and again at the
point of final layout use.

Address resolution must be mechanically equivalent to the ordinary parser:

* an explicit row `r` is XML-decoded and parsed with the checked one-based row
  grammar and must be strictly greater than the preceding row and within the
  worksheet grid;
* an omitted row `r` means the previous row plus one, with checked overflow and
  grid bounds;
* an explicit cell `r` is XML-decoded and parsed with the checked A1 grammar,
  and its row must equal the containing row;
* an omitted cell `r` means the containing row's previous column plus one, with
  checked overflow and grid bounds;
* the resulting one-based coordinates are converted to the checked zero-based
  `litchi_sheet::Cell` address without sentinels or unchecked arithmetic.

Rows must be strictly increasing. Cells within a row must be strictly
increasing by column. Every resolved cell row must equal its containing row.
Explicit and inferred row/cell forms must take the same transitions, including
empty `<row>` and `<c>` forms. Cell and row spans must cover exactly the event
boundaries consumed by the writer; `tag_end` and `close_start` must be ordered
inside their enclosing spans. `sheetData` must close before the root, and a
dimension must be a single direct worksheet child before columns or
`sheetData`, with the same checked `ref` parsing as the ordinary scanner.

The collector must make a deliberate decision about stored empty and styled
cells. The parser's `Store` includes every accepted `<c>` record, even when its
payload is empty; therefore those records participate in the same source/store
bijection and cannot be silently omitted as “not values.” Row records and
their positions must likewise remain consistent with the ordinary parser's
source order and duplicate-row checks.

## Action boundary and fallback

The collector must be attempted only for the existing compact writer boundary:
existing row/cell records and the currently supported scalar update payloads.
`Insert`, `Remove`, source `<f>` formula structures (including array,
data-table, and shared-formula forms), shared-string actions, rich inline
content, unknown or future payload variants, new rows, and any style/payload
combination outside the proven writer contract must use the complete provenance
writer. A target plain `SetFormula` remains the deliberate exception: its
existing representation is `Action::Update` with
`Payload::Set(Content::Formula)`, and it is eligible only when the source cell
is an otherwise eligible scalar record and the direct walk proves that no
source formula or uncertain child structure is present. This check must be
explicit and must be performed before relying on compact spans. A false
negative is safe; a false positive is not.

Any failure in offsets, address resolution, namespace or attribute equivalence,
ordering, cardinality, source identity, entry identity, resource accounting,
or action eligibility must drop all collector state and invoke the existing
complete rewrite with the original source and action map. The fallback must
remain the authority for exact output, typed errors, validation, readback, and
retry/error precedence. The collector must not turn a valid edit into a new
public failure solely because the optional optimization declined, and it must
not return a partially built layout.

The complete output validation, reduced-readback path, staged-value readback,
workbook calculation invalidation, source lineage/version checks, patch
construction, and clone/re-edit behavior must stay in place. A commit-local
collector is not a reason to bypass any of these checks.

## Tests required before measurement

Before a fresh candidate is applied to the restored baseline, focused tests
must make compact acceptance observable rather than allowing the wrapper's
fallback to pass vacuously. The direct collector/compact route needs exact
scanner/writer differentials for explicit and inferred rows/cells, empty and
styled cells, prefixed transitional and strict namespaces, XML entities and
CDATA, dimension and whitespace preservation, discontiguous rows, and every
accepted scalar/style action. The tests must compare output bytes, omission
spans, reduced readback, parsed addresses, values, and styles.

Separate fallback tests must cover formulas, shared formulas, shared strings,
inline rich text, removals/inserts, unsupported payloads, malformed unused
attributes, foreign/mixed namespaces, source-order and cardinality mismatches,
offset/source/Store identity mismatches, event and metadata caps, and a
populated late cap refusal that proves the ordinary parser continues to the
same result. Public integration tests must additionally cover source
lineage/version changes, cancellation, atomicity, inverse application,
multi-sheet edits, and exact publication diagnostics. At least one test must
prove that an empty or effective no-op transaction never constructs the
collector; this should be a source-level or observable counter assertion rather
than an assumption from the call site.

## Measurement and review disposition

This design has no performance result. Moving proof work from planning to
commit may improve the planning guard while increasing commit latency, commit
RSS, or transient allocations. A second ordinary semantic parse would erase
the hypothesis and must be rejected as a direct-collector result. The existing
0552 main matrix, managed path, planning guard, cap guard, quality plan,
allocation/RSS gates, exact-refusal checks, and conditional commit instruction
profile must be rerun from a freshly frozen 0553 baseline/candidate pair.

The draft-07 source gate is clear. The candidate remains pending the direct
acceptance and fallback/error-order tests, complete quality plan, matched
baseline/candidate measurements, and independent verification. If the pilot
shows a new large-tag/large-cell transient allocation or peak, the review must
assess that measured evidence before any guard is added. No source change or
gate relaxation is justified by this design note alone. ODF remains deferred
while this OLE2/OOXML optimization program continues.
