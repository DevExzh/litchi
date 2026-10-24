# 0362dvea — DrawingML chart raw review repair

- Date: 2026-10-24.
- Returned source: `96ea764602` (`feat(drawingml): bound chart raw codecs`).
- Repair source: `2b5c0906e5` (`fix(drawingml): repair chart raw admissions`).
- Status: source repair complete; coordinator-capped Cargo rerun remains
  pending because the focused build slot was released to the diagram lane.

## Review findings repaired

The returned chart raw packet had three admission gaps. ChartEx and chart-style
tree builders counted nodes for `Start` events but did not admit the actual
depth of an `Empty` event. The classic chart stream had the same latent gap in
its event adapter and fragment capture path. All three now check
`stack.len() + 1` (or the stream's current depth plus one) before retaining an
empty node, and report the typed depth dimension.

`Limits::max_attribute_value_bytes` is a decoded, normalized-value budget. The
classic reader and ChartEx parser previously compared the lexical bytes in
`Attribute::value`, which could reject an entity spelling whose decoded value
was within policy. A shared, allocation-free preflight scanner now mirrors
quick-xml's XML 1.0 entity and whitespace normalization length rules before
`decoded_and_normalized_value` can allocate. The observed decoded length is
reported for over-budget values. Chart-style attributes use the same decoded
per-attribute budget in addition to their cumulative
`max_style_string_bytes` budget; namespace declarations and ignored attributes
are covered before their owned strings are created.

The `chart::Limits` API, default ceilings, typed `Error::Limit`/
`Error::InvalidLimit` behavior, and all public raw signatures are unchanged.
DOCX/XLSX retained host policies still map into the canonical chart policy;
PPTX, spreadsheet-drawing, OPS and raw test/example consumers still select
`Limits::DEFAULT` where no narrower host policy exists. No consumer migration
or compatibility alias was introduced by this repair.

## Regression coverage

`crates/litchi-drawingml/tests/chart_limits.rs` now covers:

- exact and over-budget escaped-entity attribute values for the classic,
  ChartEx and chart-style readers, asserting decoded observations;
- an `Empty` child at the first disallowed depth for each of those readers;
- the previously existing node, depth, semantic admission, output and
  invalid-policy cases.

No Cargo command was run in this repair lane by coordinator instruction. The
changed source was reviewed with `git diff --check` and rustfmt; the prior
committed DML/consumer checks remain in 0362dve and must be rerun on this
repaired source before integration. Historical receipts and report harness
source pins were not rewritten, and no binary or large evidence artifact is
part of this record.
