# ODS sheet metadata validation

This batch addresses the ODF spreadsheet audit's missing public transactions for
consolidation, label ranges, cell-range sources, and formula-auditing (detective)
metadata. The format owner is `litchi_ods::sheet_metadata`; values reuse the
existing ODS semantic models. External sources and auditing commands remain inert.

The lifecycle target is
`crates/litchi-ods/tests/ods_sheet_metadata_transactions.rs`. Its synthetic ODF
1.4 fixtures retain their complete XML in the test source. They are not native
producer fixtures. Native acceptance and measurements, when present, have their
own receipts in this directory.

An opened spreadsheet can update inert consolidation metadata through its
ordinary facade:

```rust
use litchi_ods::{Spreadsheet, sheet_metadata::Options};

fn configure(book: &mut Spreadsheet) -> litchi_core::Result<()> {
    book.edit_sheet_metadata(|edit| {
        edit.set_consolidation(Some(Options::new(
            "sum",
            vec!["Data.A1:Data.A10".to_owned()],
            "Data.C1",
        )?))
    })
}
```

Use `sheet_metadata_with` and the explicit-context edit methods to govern
metadata parsing, staging, and candidate readback. The owned facade’s package
replacement and rehydration retain the package’s policy; they are not charged
to that metadata context. Source-backed sequential publication uses the context
supplied in `PublicationOptions`. The snapshot/edit/commit path also exposes exact
reversible patches. Cell edits accept the same logical worksheet by name or
checked position and merge their staged state before publication.

## Requirements and evidence

The frozen candidate passed **703 tests across 42 targets**, doctests, strict
Clippy, rustdoc, scoped formatting, and crate-boundary checks. The independent
source-hash verification is in [root-gates.json](root-gates.json); performance
measurements and their limits are in [report.md](report.md).

| Requirement | Executable evidence |
|---|---|
| Read all four metadata families | `metadata_snapshot_reads_four_families_and_name_position_selectors` |
| Preserve missing, empty, and covered-cell states | `metadata_snapshot_distinguishes_absent_empty_and_explicit_covered_states`, `metadata_snapshot_reports_implicit_merge_coordinates_without_projecting_anchor_metadata` |
| Create, update, remove, and invert | `metadata_edit_crud_preserves_order_and_returns_exact_inverse_patch`, `metadata_edit_creates_and_clears_absent_singletons_without_collapsing_empty_labels` |
| Split only selected logical occurrences in repeated rows/cells | `metadata_edit_splits_repeated_row_and_cell_runs_for_one_logical_target`, `metadata_edit_merges_two_targets_in_one_repeated_row_and_cell_run` |
| Exact no-op and revert | `metadata_edit_reverts_source_and_detective_to_exact_noop`, `metadata_noop_on_repeated_runs_retains_exact_source_without_splitting` |
| Preserve namespaces and unrelated XML | `metadata_scanner_uses_namespace_uris_and_keeps_same_named_foreign_wrappers_opaque`, `changed_owner_preserves_comments_processing_instructions_and_foreign_siblings` |
| Reject malformed or ambiguous ownership | `metadata_scanner_refuses_malformed_direct_owner_order_and_duplicate_owners_before_staging`, `metadata_scanner_rejects_unbound_namespaces_duplicate_attributes_and_bad_values` |
| Retain charged memory through patch lifetime | `managed_patch_retains_source_and_target_memory_until_final_patch_drop` |
| Cancellation and source freshness | `metadata_patch_apply_honors_retained_context_cancellation`, `source_backed_metadata_detects_stale_source_before_publication` |
| Sequential package publication and unchanged members | `source_backed_metadata_commit_round_trips_inverse_and_untouched_members` |
| Signed-source refusal | `signed_source_refuses_changed_metadata_while_preserving_staged_edit` |
| Ordinary and mutable facade integration | `ordinary_and_mutable_facades_forward_metadata_snapshot_edit_and_patch` |
| Name and position select one staged identity | `metadata_edit_canonicalizes_name_and_position_selectors_to_one_cell`, `metadata_edit_mixed_selector_revert_returns_exact_noop` |
| Standard ODF root siblings survive | `metadata_scanner_accepts_standard_document_content_siblings` |
| Owned and mutable signed-source refusal | `ordinary_signed_facades_refuse_changed_metadata_publication` |

The separate `tests/sheet_metadata_context.rs` target checks that equal budget
limits, usage, and diagnostic scope strings cannot impersonate resource
ownership; cloned contexts retain authority; and a fresh token on a shared budget
cannot bypass cancellation of the source context, including exact no-ops.
It also checks destination output limits for ordinary patches and for a second
metadata profile on the same source-backed package owner. Both destination-limit
tests reproduced the original bypass before the admission fix. Additional
assertions require structured `ResourceLimit` diagnostics for both scanners’
input ceilings, staged operation counts, and patch destination output ceilings.

`sheet_metadata_xml_conformance.rs` adds paired-empty owners, root family/order
checks, prelude insertion, and namespace scope regressions. See [review.md](review.md)
for the review corrections and their boundaries.

[validate_native_owners.py](validate_native_owners.py) uses Python and `lxml` to
select each owner's definition from the bundled ODF 1.4 RNG while retaining its
datatype dependencies. [native-owner-schema.json](native-owner-schema.json)
records the exact native-input package and schema hashes. Owner validation and
whole-content validation are reported separately. Replay with `--spec
3rdparty/specs/OpenDocument-v1.4-os.zip --package <input.ods>`.

These mappings identify test scope; final command exits and source hashes are
required to establish that a particular candidate passed. Review corrections
and additional regression coverage are recorded with the final gate receipt.

## Boundaries

This owner edits stored metadata. It does not calculate consolidation results,
apply scenarios, calculate dependencies, draw auditing arrows, refresh links,
or fetch external resources. Exact in-memory patches do not establish durable
patch exchange or concurrent edit merging. The wider specification audit and
the end-to-end performance program remain open.
