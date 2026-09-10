# XLSB worksheet binary-index review

Date: 2026-09-10

Scope: read-only review of the current `litchi-xlsb` worksheet binary-index
codec, source-backed cached-value handle, writer emission, and worksheet
publication hooks. This review covers the current worksheet-index batch only;
it does not review unrelated XLSB projections.

## Normative basis

The checked-in `[MS-XLSB]` material identifies the binary-index content type
and relationship at `3rdparty/specs/[MS-XLSB]/2 Structures/2.1 File Structure.md:1543-1557`.
That section says what an existing Worksheet Binary Index part must look like
in the OPC graph; it does not say that every worksheet must contain one.

The local conceptual contract is at
`3rdparty/specs/[MS-XLSB]/2 Structures/2.2 Conceptual Overview.md:19-35`:
lookup selects a `BrtIndexBlock`, uses its following optional
`BrtIndexRowBlock`, then scans the worksheet from the computed cell-record
offset. A missing matching block means the row has no data or formatting.
The same section at lines 11-17 defines a non-empty row to include row or cell
formatting, and enumerates the indexed cell record kinds.

The record contracts are at
`3rdparty/specs/[MS-XLSB]/2 Structures/2.4 Records.md:39503-39613`
(`BrtIndexBlock`), `39711-39723` (`BrtIndexPartEnd`), and
`39715-39838` (`BrtIndexRowBlock`). In particular, the row mask's column-mask
array has one item per set row, and the sub-offset array has one item per set
column-block bit. The `unused1`, `unused2`, and `unused3` fields of
`BrtIndexBlock` are undefined and MUST be ignored. The checked-in Markdown
gives the external grammar filename but not its rule body. The exact rule used
here is from the official Microsoft `[MS-XLSB]` v20250916 PDF, page 93:
<https://officeprotocoldoc.z19.web.core.windows.net/files/MS-XLSB/%5BMS-XLSB%5D-250916.pdf>.
It is:

```text
SHEETINDEX = 1*(BrtIndexBlock [BrtIndexRowBlock]) 1*2BrtIndexPartEnd
```

This is the basis for the local parser's required block, optional row block,
and one or two end-marker handling. The numeric cached-value rule is
`3rdparty/specs/[MS-XLSB]/2 Structures/2.5 Structures.md:26059-26061`:
`Xnum` rejects infinity, denormals, NaN, and negative zero. Rich-string
limits and ignored flag bits are specified at lines 22863-22879.

## Confirmed implementation behavior

The current codec correctly handles the important grammar and preservation
cases:

* It requires at least one `BrtIndexBlock`, accepts an absent row block,
  accepts one or two `BrtIndexPartEnd` records, and rejects records outside
  `SHEETINDEX`.
* It validates the row-range ceiling and exact variable-field lengths,
  checks overlap and duplicate anchors, and binds every modeled offset to an
  actual worksheet cell-record boundary.
* It accepts zero column-mask bits for a set row. The native
  `testVarious.xlsb` fixture contains formatting-only row headers with this
  shape; the encoder now retains those row headers and the integration test
  verifies that topology regeneration does not drop them.
* It applies the exact Xnum rule to RK, real, and formula-number caches. Inline
  and formula strings enforce the 32,767-character cell ceiling. Rich strings
  are bounded before the full rich/phonetic semantic parser runs, while
  undefined RichStr flag bits remain ignored.
* The source-backed handle materializes and retains the worksheet `PartData`,
  checks source freshness and cancellation before and after lookups, verifies
  declared part sizes before `data()`, and applies the caller's string ceiling
  to a returned shared-string copy. Formula tokens remain intentionally
  opaque; only the stored cached scalar is exposed.
* Worksheet edits stage a candidate package before replacing the worksheet.
  Cell-value, scenario, timeline, slicer, sparkline, drawing-transfer, and XML
  Maps paths route worksheet changes through the shared index maintenance
  helper before publication. A malformed index that must be reparsed leaves
  the candidate unpublished. The writer creates a valid empty block for an
  empty worksheet and generates offsets only after finalized worksheet bytes.

## Findings and disposition

### 1. Index-copy ceiling — resolved

`binary_index::maintain_index` now checks the source-index byte ceiling before
either the exact-no-op or geometry-stable preservation copy, using the same
`min(MAX_INDEX_STREAM_BYTES, limits.source_bytes)` ceiling as parsing. The
regression `preservation_copy_is_bounded_before_noop_or_geometry_scan` covers
both paths.

### 2. Publication record ceiling — resolved

`Limits::publication_default` now derives `max_records` as
`worksheet.source_bytes() / 2`, with an explicit two-byte minimum BIFF12
record-framing assumption. This is a valid inclusive upper bound for every
record stream within the existing worksheet byte ceiling, and it preserves the
cell owner's established byte policy without introducing a hidden smaller
limit. The code comment documents that relationship.

### 3. Geometry-stable opaque preservation is acceptable only as an explicit
bounded preserve boundary

The current fast path intentionally preserves unknown or malformed index bytes
for a non-byte-identical worksheet edit when the scanned row-header list and
compact first-cell bucket geometry are unchanged. This is compatible with the
lossless-preservation policy if it is described as a source-preserving
exception: publication does not interpret or repair that index, and the
source-backed lookup API still parses and binds it strictly, so a malformed or
unknown index remains an error at lookup time. Unknown extension records may
carry semantics outside the modeled grammar; the implementation cannot claim
to update those semantics.

Documentation action: state this condition precisely in the maintenance and
feature-matrix text, and avoid a blanket claim that every stale or unsupported
index is refused on every worksheet mutation. The
offset-changing and row/bucket-topology-changing paths still parse and bind
before patching or canonical regeneration and therefore refuse unknown index
records through the exact `SHEETINDEX` parser.

### 4. Small normative-profile limitations should be named

The parser additionally requires strictly increasing, non-overlapping index
block ranges and rejects zero-length ranges. The cited `BrtIndexBlock` field
rules specify the upper range and maximum row bounds but do not state an
ordering production or an explicit `rwMac > rwMic` requirement. Refusing such
inputs can be a reasonable conservative producer profile, but the public
documentation should call it a bounded refusal rather than imply that the
local ABNF alone requires it.

The writer comment at `crates/litchi-xlsb/src/writer/workbook/package.rs`
currently says every worksheet MUST have a binary-index part. The normative
file-structure section requires the relationship and content type for an
existing index part; it does not impose that universal worksheet requirement.
The writer may continue to emit an index for all authored worksheets, but the
comment should describe that as the writer's chosen output profile.

## Validation performed

The focused current snapshot passed:

* `cargo check -p litchi-xlsb --lib`
* 11 `binary_index` unit tests, including Xnum, string, empty-block,
  non-aligned partition, and geometry-stable opaque-preservation cases
* 18 `worksheet_binary_index` integration tests, including source limits,
  retained `PartData`, stale/cancellation refusal, formatting-only rows,
  offset patching, topology regeneration, and atomic opaque-index refusal

## Disposition

**Code review clear.** The two resource-bound findings are fixed and covered by
focused tests. The core grammar, source binding, scalar decoder, worksheet
publication routing, and formatting-row preservation are sound in the reviewed
scope. Geometry-stable opaque preservation is an accepted bounded
lossless-preservation boundary; its public maintenance/feature-matrix wording
must describe that exception accurately. The strict ordering/nonempty-range
checks remain a bounded conservative producer profile, not a claim about an
ABNF ordering production.
