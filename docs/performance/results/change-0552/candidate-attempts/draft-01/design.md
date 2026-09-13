# Change-0552 draft-01: compact source-cell proof

## Status and boundary

This note describes the frozen first compile/correctness draft captured in
`inputs.json`. It is an implementation record, not an admission or a
performance result. The base revision is
`3d84cf64cd3ffbbd809e8677f37c331e93cf0b8c`; the frozen patch SHA-256 is
`312859df0623c868354037940459498e24c80384ea4eac2fac49c74ce201688e`.

The draft targets source-backed XLSX `MultiSourceEdit`. It does not change
OLE2 or ODF code and does not change the single-sheet `SourceEdit` route.
The complete scanner/writer remains the behavior oracle and fallback.

## Traversal and proof representation

An eligible worksheet is traversed once by the existing namespace-aware raw
parser and the source validator. The traversal records each reader event's
original start and end positions and decoder. The validator callback remains
pre-transition. A separate post-transition handoff reports parser-resolved
row and cell addresses for row/cell Start, Empty, and End events. Empty cells
and empty rows therefore receive their own event-timed handoff; no address is
recovered from a preceding materialized cell.

`CompactProofBuilder` retains only source-bound envelopes and offsets:

* each cell has two `u32` offsets (`CompactSpan`, 8 bytes);
* each row retains its number, row span, opening/closing boundaries, and a
  boxed slice of compact cell spans;
* sheetData retains its span, opening/closing boundaries, row slices, and its
  empty form; and
* an optional dimension span retains its declared `Rect` and Start/Empty
  form.

The builder decodes every opening Start/Empty element name and every
attribute name and value, including namespace declarations and attributes
that are not later copied. Values use the scanner's XML 1.0
decode/normalization operation. It checks root, sheetData, defaults, columns,
dimension, row/cell order, duplicate/order constraints, inferred addresses,
empty forms, scalar primary cardinality, source bounds, and row/cell
cardinality. Formula elements deliberately disable the proof so formula
range and shared-group behavior stays with the complete scanner.

The compact plan zips the compact spans with the exact `Store::entries()`
slice published by the same source parse. It requires strict row/column
ordering, equal cardinality, and an exact entries allocation pointer/length.
This avoids retaining one address object per cell while preventing a foreign
store with coincident counts from authorizing source offsets.

## Resource limits and refusal behavior

The source traversal keeps the existing eligibility limits: an 8 MiB source,
the MCE/x14ac marker exclusions, UTF-8 input, and the shared 131,072-event
bound. The proof's retained metadata budget is 2 MiB. Fixed builder/layout
metadata is charged up front. Stack, row, and cell vectors use fallible
checked geometric growth; the requested growth and actual capacity delta are
charged before the reserve. A failed charge, reserve, conversion, or proof
invariant drops all builder scratch and publishes `None`.

Proof refusal does not return false from the existing validator observer and
does not request source parser replay. The parser and validator continue to
their established result. Reader, validator, parser, and raw-source errors
therefore follow the pre-existing fallback sequence. If the proof is absent,
or if source/store identity, bounds, or compact-plan eligibility fails at
commit, the complete value-only scanner/writer is used.

The 2 MiB accounting in this draft is a retained proof metadata bound. It is
not yet a complete bound over all temporary allocations made by the shared
validator/parser or by the compact writer. In particular, a changed cell can
materialize decoded attribute names and values through `cell_tag`, a changed
dimension can materialize a `Tag`, and the primary span can be boxed. Those
writer-side allocations are currently fallible only at the `Result`-returning
decode boundaries; their temporary size is not preflighted against the proof
cap. This is the provisional review hold and must be resolved or measured
before admission. The existing complete scanner/writer's allocation behavior
is inherited baseline behavior; the compact route must not introduce a new
unbounded materialization boundary while claiming the proof cap covers it.

## Compact writing and retained readback

For an eligible existing-cell scalar update, the compact writer copies the
root, pre-data bytes, unchanged rows, and unchanged cells directly from the
immutable source. It lazily reads only a changed cell's span to reconstruct
the ordinary ephemeral cell tag and direct scalar primary span, then invokes
the existing cell writer. Dimension expansion, when needed, lazily reads the
dimension opening event and uses the declared range unioned with every
retained non-removed store entry. The writer records output omission spans in
the same form as the complete value-only writer.

The compact path rejects removal, missing-cell membership, and shared-formula
actions; those actions use the complete writer. The current crate-private
eligibility predicate explicitly rejects `Remove` and shared-formula payloads
and matches the public source-backed action set, but it should be tightened to
an explicit supported-payload list before this becomes a stable internal
boundary. In particular, an internal caller supplying `SharedString` or a
future payload variant should not be admitted accidentally.

Every compact output still passes worksheet XML validation. Snapshot rebinding
uses the existing reduced independent readback only when omission metadata is
valid; otherwise it parses the complete output. Complete bytes, omission
spans, and reduced semantic records remain the differential oracle.

## Layout fields and fallback obligations

The strict shared validator admits only the dependency-free worksheet grammar.
It proves the direct SpreadsheetML root and sheetData envelope, checks
defaults/columns/dimension ordering and full attribute decoding, and rejects
protection, merge, validation, compatibility, MCE, x14ac, and unknown
dependency-bearing elements. Those rejected forms remain on the complete
route. Formula ranges and shared-formula state are not inferred absent;
formula sources disable compact proof. `merge_insertion` is not needed by the
existing-cell scalar writer and is never fabricated for the compact route.

The proof retains no full `Layout`, scanner arena, row-only parse, or per-cell
owned tags/primary vectors. It relies on the already completed validator and
raw parser for all semantic errors, then uses only the compact facts and
ephemeral changed-cell views required by the existing writer. Any uncertainty
in those facts declines the proof and leaves the authoritative full scanner
and writer responsible for diagnostics and output.

## Ownership and invalidation

`CompactLayout` is immutable and contains source allocation pointer/length and
the parser Store entries allocation pointer/length. The compact writer only
compares these identities; it never dereferences a stored pointer. The
Snapshot owns the matching immutable `SourcePayload` and `Arc<Store>`, so the
allocations remain pinned for the proof's lifetime. Snapshot clones share the
same ownership. Source lineage/version checks continue to guard patch replay,
and every worksheet payload replacement clears the compact proof. Workbook
only invalidation leaves the worksheet allocation and proof unchanged. The
layout's fixed fields and boxed immutable slices are suitable for immutable
sharing; no mutable proof state escapes the builder.

## Verification still required

The private differential module covers accepted and refused prefixed, typed,
inferred, empty, styled, CDATA, dimension, formula, source-identity,
Store-identity, malformed-attribute, and cap cases, including bytes, omission
spans, and reduced readback. The independent public guards cover the
baseline-compatible integration cases. The frozen draft still requires the
targeted build, the broader exact error/output oracle, writer allocation
refusal coverage, and the native/allocator workflow, no-op memory,
managed-execution, and retained-memory gates from the 0552 plan. No
performance claim or admission decision follows from this draft alone.


The first `cargo check --locked -p litchi-xlsx --all-features` was run by the
root owner after this source snapshot was frozen and exited 101. Its receipt
is `docs/performance/results/change-0552/check-attempts/candidate-compile-01/receipt.json`.
The diagnostics are mechanical integration issues in this draft (private
re-export visibility, unused qualification, a `return false` in a unit
function, an overly restrictive `Sized` inference, boxed-slice type
inference, and a trivial cast); no compile-success or runtime-correctness
claim is made here. The exact draft remains preserved for that review, and
follow-up source edits belong to a separately frozen attempt.
