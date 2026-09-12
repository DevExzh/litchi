# Next OLE2/OOXML priority

This is a bounded source and profiling audit after the 0539 private XLSX
transient-Cow pilot. The pilot remains rejected by native results: total p50
and mean fail in three of four primary rows, and one of four planning p50
rows fails. The passing planning-allocation gate is a separate diagnostic and
does not change native admission. Do not relax the gates or infer speedup from
allocation-call counts.

## Selected worksheet validation/parser boundary

The next target is a proof-bearing audit of the selected worksheet load path,
owned by
`SourceBackedEditor::edit_sheets` → `MultiSnapshot::load_source_backed` →
`Snapshot::from_source_selected` in
[`snapshot.rs`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L426).
The exact sibling stages are `cell_values::validation::worksheet_xml` →
`validation::validate_xml` and `raw::worksheet::parse` →
`raw::worksheet::codec::Parser::parse`; the latter also performs MCE handling
and the x14ac extension scan in [`raw/worksheet/mod.rs`](../../../../crates/litchi-xlsx/src/raw/worksheet/mod.rs#L28).

Historical 0530 selected `edit_sheets` profiles make this a high-return audit
target:

| profile | raw worksheet parse | XML validation |
| --- | ---: | ---: |
| medium | 75,518,721 (60.1%) | 46,622,053 (37.1%) |
| dense-sparse | 141,675,512 (59.8%) | 89,730,098 (37.9%) |

Both stages consume the complete worksheet byte stream and independently
create/read XML events and resolve namespaces. That is evidence of duplicated
traversal, not evidence that either stage can be removed: validation currently
establishes the error boundary before raw parsing, while raw parsing performs
MCE/x14ac processing and constructs cell semantics.

Before editing, measure fresh release planning runs with exact symbol and call
edges for `worksheet_xml`, `validate_xml`, `raw::worksheet::parse`,
`Parser::parse`, `process_ooxml`, and the x14ac scan. Keep medium and
dense-sparse profiles plus the five XLSX guards, and record stage-local and
whole-plan native p50/p95/p99, mean, RSS, and allocation counts/bytes. Use the
same ABBA repeats, source/output/cache identities, budgets, and execution
fences as the primary verifier; keep commit and publication allocation metrics
separate. Counts must not be converted into a latency claim.

The differential fixture set must cover ordinary, transitional, and strict
worksheets; MCE and foreign namespaces; x14ac extensions; unknown, qualified,
duplicate, and malformed attributes; DTD/PI, deep/truncated, mismatched-end,
text, and general-reference errors; formulas and shared formulas; and rows,
columns, merges, styles, metadata, and source retention. Capture the first
error domain and order as well as parsed/output identity. Any candidate must
retain dialect/root/depth/closing checks, full checked attribute iteration,
MCE and x14ac typed capture, the historical repeat-on-parser-error behavior,
semantic materialization, style and metadata checks, resource/cancellation
and execution/version fences, and aggregate bounds. A private validated
worksheet handoff is worth considering only if these measurements prove a
reusable result with those obligations; do not bypass or fuse the stages on
historical percentages alone. ODF work remains deferred until the OLE2/OOXML
optimization goal is complete.
