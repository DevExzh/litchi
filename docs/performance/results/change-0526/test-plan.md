# 0526 row-primary-span arena test plan

This is a bounded correctness plan for the proposed XLSX row-owned primary
span arena. It is a test plan only; the candidate production change is not
present in this worktree. OLE2/OOXML remains the active optimization lane and
ODF remains deferred.

## Review base

The plan is bound to accepted commit
`67028ab6037ae6eef15af92a0d540285c3c5362c` and the unchanged source recorded
in `source-binding.json`. The un-applied draft is
`row-primary-arena.patch`; its SHA-256 is
`28269e6051f9ee9453ccfd590a0f89269c400529d973a84f0ac908eb6730c5e9`. The
relevant baseline and draft source SHA-256 values are:

| File | Base SHA-256 | Draft SHA-256 |
| --- | --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs` | `71e23d40e8ef116ae833b51fb77c11067461fc6d5fb69f05fb2bfbe61644a0bd` | `fea4369470a24600451bc96e7992e8a346b1c61a689ebc57d8e400da001ce736` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8` | `883ee4dcd9d848185401278665c631970712453bb6b5a73010228da3204a96a1` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5` | `4d9841a90d7b76f946bc33f394abb9f35586179952873e9447c19c7d79a64064` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/tests.rs` | `fd6fa40c037533a9546a74816202955adb18063e4b3d9f1c08323c1466f13879` | not in draft |
| `crates/litchi-xlsx/src/raw/worksheet/edit/tests.rs` | `aa7aaffa69b0ccf3f3426776a6c5ed48c23a9154b03b1ce29d7d3116e30b0228` | not in draft |

Refresh this table and the binding after the candidate is applied; the draft
values identify an un-applied review artifact and do not establish a retained
production change.

## Existing regression floor

Keep these exact tests. They cover the surrounding contracts and should remain
the oracle when the storage representation changes.

The private scanner and tag tests in
`crates/litchi-xlsx/src/raw/worksheet/edit/codec/tests.rs` are:

- `snapshot_scan_builds_sorted_edit_slots`
- `snapshot_scan_restores_namespace_scope_after_a_rebinding`
- `snapshot_scan_marks_only_plain_cells_as_tagless`
- `snapshot_scan_keeps_legacy_cell_reference_and_tag_semantics_together`
- `snapshot_scan_rejects_mismatched_end_name`
- `snapshot_scan_preserves_cell_error_precedence_through_trailing_attributes`
- `snapshot_scan_rejects_a_truncated_cell_start_after_reference_attributes`
- `snapshot_scan_reports_cell_address_errors_before_compact_tag_errors`
- `compact_cell_tag_keeps_attribute_error_order_and_messages`
- `snapshot_scan_rejects_nesting_beyond_worksheet_depth_limit`
- `snapshot_scan_accepts_nesting_at_worksheet_depth_limit`
- `snapshot_scan_handles_large_flat_event_stream_within_depth_limit`
- `snapshot_scan_accepts_event_count_at_limit`
- `snapshot_scan_rejects_flat_event_stream_over_event_limit`

The ordinary writer tests in `crates/litchi-xlsx/src/raw/worksheet/edit/tests.rs`
are:

- `untouched_xml_fragments_remain_byte_exact`
- `minimally_rewrites_set_clear_remove_and_new_rows`
- `dimension_expansion_never_narrows_producer_bounds`
- `row_visibility_surgery_is_sparse_lossless_and_composes_with_cells`
- `row_layout_facets_preserve_unedited_state_and_materialize_sparsely`
- `worksheet_defaults_and_row_descent_rewrite_losslessly_by_facet`
- `new_descent_injects_collision_free_ignorable_namespaces`
- `style_effects_preserve_payload_and_compose_with_value_effects`
- `blocks_dependencies_instead_of_guessing`

The accepted 0525 source-backed/readback tests in
`crates/litchi-xlsx/src/cell_values/snapshot.rs` are:

- `provenance_writer_matches_ordinary_rewrite_bytes`
- `merge_omitted_cells_handles_same_row_gaps_and_a_lower_next_row_column`
- `replacement_only_readback_matches_full_parse_for_all_scalar_facets`
- `replacement_after_implicit_cells_preserves_every_omitted_address`
- `changed_output_bytes_are_observed_by_specialized_readback`
- `omission_proof_from_another_source_cannot_reuse_its_store`
- `clear_remove_and_existing_row_insert_match_the_complete_parser`
- `new_row_and_shared_formula_inputs_use_complete_parse_semantics`
- `malformed_or_mce_candidate_bytes_are_rejected_before_merge`
- `no_op_source_edit_shares_the_original_semantic_store`

The new tests below target the representation-specific failure modes that the
floor does not expose. They should not replace any of the tests above.

## Minimal additions

### 1. Direct row-arena range proof

Add a private codec test, preferably named
`snapshot_scan_indexes_row_primary_spans_without_cross_owner_ranges`, to the
existing `codec/tests.rs` module (or a `#[cfg(test)]` child module with access
to the same private model). Use one worksheet fixture with all of these
features:

- row 1, cell A1 has three direct main-namespace primary elements in source
  order, for example `<v>`, `<f>`, and `<is>`, with a comment and an
  `extLst` containing a future-namespace element between them;
- row 1 also has a one-primary B1 and an empty/self-closing C1;
- row 2 is an empty/self-closing row;
- a later row uses a different prefix or nested default namespace bound to the
  SpreadsheetML namespace and has repeated primary elements with markers
  distinct from row 1.

The test may use a candidate crate-private accessor for the row arena and each
cell's range; it must assert behavior rather than vector capacity or a
particular range type. Resolve every cell range through its owning row and
assert that:

- the flattened arena count equals the sum of the primary counts in that row;
- every resolved span is sorted, non-overlapping, inside its cell's
  `tag_end..close_start`, and slices the exact expected source bytes;
- adjacent cells and rows never resolve one another's marker bytes;
- the later row's ranges resolve its own payloads even when their local indexes
  match row 1;
- empty C1 and the empty row have no primary range and do not create phantom
  arena entries; and
- comments, `extLst`, namespace declarations, and whitespace between primary
  spans are outside the removed spans and remain observable in the source
  slice.

This catches a cell storing only the last span, an off-by-one end, a global
arena index used with a row-local arena, and a range copied from a neighboring
cell. Counts alone are insufficient: compare unique marker bytes for every
span. This is intentionally a scanner/layout test over lossless XML; do not
feed its repeated/conflicting primary payloads to `Snapshot::load` or claim
semantic-parser acceptance for this fixture.

### 2. Unknown/MCE boundary after multiple primary spans

Add a private codec/edit test named
`snapshot_scan_records_interleaved_primary_spans_before_mce_refusal`. Make a
cell contain a primary element, a direct unknown future-namespace child or
`mc:AlternateContent`, another primary element, and an `extLst` child. Assert
that scanning retains all primary spans in source order and sets the existing
markup-compatibility flag. An attempted value replacement must return the
existing `EditBlock::MarkupCompatibility` error without publishing output.

This is deliberately separate from the eligible writer case: a direct unknown
child is a refusal boundary, while an unknown element nested under `extLst`
is interleaved markup that must remain in an eligible cell. The test catches a
scanner that stops collecting spans when it sees the unknown child or a writer
that silently edits around an MCE payload. Keep
`blocks_dependencies_instead_of_guessing` and
`malformed_or_mce_candidate_bytes_are_rejected_before_merge` as the broader
refusal/readback checks. This case is scanner/error-boundary coverage only;
the deliberately unknown payload is not a source-backed semantic fixture.

### 3. Raw ordinary/provenance byte differential with arbitrary spans

Add `ordinary_and_provenance_writers_match_for_nonsemantic_multi_primary_bytes`
to the raw edit tests (or a private package test), calling `rewrite` and
`rewrite_value_only_with_provenance` directly. Use a deliberately lossless but
nonsemantic cell with repeated direct `<v>`, `<f>`, and/or `<is>` elements,
comments, whitespace, and an `extLst` future element, plus an unchanged owner
in the same row and an untouched row. Do not route this fixture through
`Snapshot::load`, the source-backed full validator, or the semantic worksheet
parser: repeated/conflicting primary elements and opaque markup are the reason
this fixture exists.

For the identical action map, assert all of the following:

- `rewrite_value_only_with_provenance(...).bytes` equals ordinary `rewrite`
  byte-for-byte;
- `omitted` is nonempty and includes an unchanged owner in the touched row,
  so a blanket complete-parse fallback cannot satisfy the test;
- each omitted output slice is byte-identical to its corresponding unchanged
  source cell/row record, including its prefix, aliases, attributes, comments,
  and whitespace;
- the changed output has one explicit `r` attribute, preserves every
  non-primary interleaved byte, and removes all old primary spans rather than
  only the last one; and
- changed-cell output and omission addresses are checked against the scanned
  source spans, without making a semantic acceptance claim for the fixture.

This is the raw writer proof for arbitrary multi-primary ranges. The existing
`provenance_writer_matches_ordinary_rewrite_bytes` test remains the valid
scalar/source-backed differential; the 0526 case must exercise every span and
the bytes between those spans.

### 4. Valid source-backed provenance differential with forced omission

Add `valid_formula_value_provenance_forces_omission_and_matches_complete_parse`
to `crates/litchi-xlsx/src/cell_values/snapshot.rs`, using its compact package
helpers and `assert_store_matches`. Use a source-backed-valid worksheet with a
single normal formula/cache pair (`<f>...</f><v>...</v>`), comments and
prefixed/default SpreadsheetML aliases, an unchanged B1 in the touched row,
and an untouched row. Do not include `extLst`, direct unknown children, or
repeated/conflicting primaries. Change only A1 so the value-only route is
eligible and omission is forced.

Assert that the proof is eligible (`omitted` is nonempty and includes the
untouched owner), that its bytes equal ordinary `rewrite` for the same action,
and that a complete parse of the exact emitted bytes equals specialized
readback across values, styles, row/column/default entries, declared and
effective extents, formula fields, and shared-formula metadata. Assert the
changed A1 has an explicit `r` and that the omitted B1/row bytes remain exact.
Use `changed_output_bytes_are_observed_by_specialized_readback` for the
mutation seam rather than replaying a staged value. This is the semantic
oracle test; it is intentionally separate from test 3's arbitrary raw XML.

### 5. Style-only and empty-owner writer path

Add `style_only_multi_primary_and_empty_row_rewrite_is_lossless` to
`crates/litchi-xlsx/src/raw/worksheet/edit/tests.rs`. Use a row containing a
multi-primary A1 with an `extLst` and comment, an empty/self-closing B1, and a
self-closing empty row immediately after it. Apply a style-only action to A1
and a style materialization or reset to B1.

Assert that ordinary rewrite changes only the intended cell start tag; all A1
primary payloads and interleaved bytes are preserved exactly, B1 remains a
valid empty cell with the expected style, and the empty row remains an exact
self-closing row. Compare the raw output to the expected start-tag and source
slices only; do not feed this deliberately duplicate-primary fixture to
`Snapshot::load` or the semantic parser. This specifically exercises the
writer branch that has no replacement payload and therefore must not resolve a
bogus arena range. Keep `style_effects_preserve_payload_and_compose_with_value_effects`
as the valid semantic style/value composition test. If an additional semantic
style assertion is needed, use a separate valid fixture with one `<f>+<v>`
pair and no opaque extension payloads.

### 6. Malformed/error-order after arena activity

Add `snapshot_scan_preserves_malformed_error_order_after_prior_primary_spans`
to `codec/tests.rs`. Prefix each malformed case with a valid row containing
several primary spans, then run the existing malformed fixtures for:

- an invalid cell reference combined with a malformed/unsupported cell tag;
- a truncated cell start after reference attributes; and
- a mismatched closing name or truncated primary/cell close.

Compare the first two cases with `legacy_cell_pipeline_error` and retain the
existing exact error strings; compare the third with the current scanner
structure error. The populated prior row must not change the error, introduce
an arena-bound error, or lose the prior row. This extends
`snapshot_scan_preserves_cell_error_precedence_through_trailing_attributes`,
`snapshot_scan_rejects_a_truncated_cell_start_after_reference_attributes`, and
`snapshot_scan_reports_cell_address_errors_before_compact_tag_errors` to the
stateful arena boundary.

### 7. Checked invalid-range refusal

Add one narrow private writer test, preferably named
`write_cell_rejects_primary_range_outside_owning_row`, beside `write_cell` in
`codec/snapshot/write/sheet_data.rs` (or in a test-only child module with
access to that private function). Start from a valid scanned cell, clone its
slot, and replace its range with a half-open range whose end exceeds the
owning row arena. Invoke the payload-replacement branch directly and assert a
typed `Error::Invalid` with the draft guard's exact message:
`worksheet cell primary span range exceeds its owning row`.

The assertion must prove a malformed range returns an error without panicking;
the package-level caller must discard its local output on that error. Do not
add a broad synthetic layout matrix here. This single test protects the new
`primary.get(cell.primary.clone())` check from becoming an unchecked index or
silently falling back to another row's arena.

## Oracle and execution boundary

The direct codec tests should inspect source spans and scanner flags. The raw
multi-primary and style-only fixtures should use ordinary raw rewrite bytes,
direct provenance bytes, and exact source slices as their oracles; they must
not claim acceptance by `Snapshot::load` or the semantic parser. Only the
valid source-backed fixture in test 4 should use a fresh full raw worksheet
parse of the exact emitted bytes as its semantic oracle. There, a candidate
readback or retained store is correct only when it equals that full parse
across values, styles, rows, columns, defaults, extents, formula and
shared-formula metadata. Assert eligibility (`omitted` is nonempty) in every
test that is intended to exercise reuse; assert empty provenance and complete
parse behavior for new-row and shared-formula fallback using the existing
`new_row_and_shared_formula_inputs_use_complete_parse_semantics` test.

No build, test, capture, or Rust source edit was performed while preparing this
plan. The parent/root agent owns the candidate application and all validation
lanes.
