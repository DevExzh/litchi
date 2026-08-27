# Change 0330: XLSB source-backed table context

## Decision

Source-backed XLSB catalog opening now validates the metadata needed to
identify internal `TABLE` and `STRICT_TABLE` relationships, including their
content types, without reading any table payload. The catalog preserves exact
worksheet positions, so later selection remains tied to the source workbook's
sheet order rather than to a reconstructed or normalized view.

The first selected worksheet materialization lazily parses all workbook table
definitions into one immutable, retryable semantic cache. This cache is shared
by subsequent worksheet materializations and supplies the workbook context
needed for cross-sheet structured references. A failed attempt does not publish
partial context: cancellation or source mutation observed while processing the
second table aborts publication, and a later retry rebuilds the complete
context. Semantic reuse was verified with OPC zero-retention instrumentation.

## Preservation and errors

Formula text and source-cached values remain preserved independently of table
context resolution. Catalog errors are typed for wrong table content types and
external relationships. Materialization reports typed errors for malformed
table metadata and duplicate table identifiers or case-insensitive table
names. External tables and `PivotTable` relationships remain explicitly
refused; this change does not fetch external targets or infer pivot scope.

The implementation preserves lossless opaque content while making the table
context boundary explicit. A table definition is not treated as successfully
resolved merely because its relationship was cataloged.

## Validation evidence

- `542` XLSB library tests passed.
- `121` integration tests passed, with the exact persistent drawing test
  `checked_in_unique_standard_drawing_corpus_transfers_every_anchor` skipped
  as the known unrelated failure.
- Focused source-backed coverage: `31` tests passed.
- Strict Clippy and rustdoc checks passed, together with the facade feature
  check, formatting check, and `git diff --check`.
- Validation used serialized `CARGO_BUILD_JOBS=1` execution and one isolated
  target directory: `/dev/shm/litchi-0330-target`.

The isolated target was deleted after validation.

## Nonclaims and boundaries

This change makes no claim of external-target fetching, `PivotTable` scope
resolution, formula evaluation, or recalculation. It makes no performance,
RSS, OOM, or complete semantic-memory-bound claim. `ReadAt` provides no atomic
guarantee if the source mutates after the final version observation. A
late-invalidated `OnceCell` allocation may remain retained, but it cannot be
returned without passing the source and execution fences.
