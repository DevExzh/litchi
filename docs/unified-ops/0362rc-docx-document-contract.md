# 0362rc — DOCX document owner and raw Tree contract

This change closes the DOCX main-document surface left by the 0362 review.
`document::raw::Tree` is the package-neutral mutable WordprocessingML codec.
It admits XML through required `document::raw::Limits`, including caller
selected byte, depth, element, per-element attribute, attribute-byte,
operation, and replacement-text bounds. The parser checks XML structure before
retaining its layout, and hard format ceilings remain separate from the caller
policy. Output and replay may tighten the retained policy but cannot widen it;
failed mutations leave the projected tree unchanged and repeated `&self`
output copies only the current XML bytes.

Raw trees expose reversible `document::Patch` planning and replay without a
raw snapshot-publishing commit. Package owners publish through the canonical
`document::Commit<P>` triad. `document::composition::{Prepared, Composition,
History}` retain recoverable joins, three-way plans, and bounded history;
`Package::apply_composition` and `source_backed::Package::apply_composition`
stage all accepted operations in one owner edit and commit. Owner-created
compositions carry a complete artifact fingerprint, so an independently
reopened equal-content package is accepted while a package with a changed
untouched part is refused before document staging. A composition made directly
from a raw XML tree retains XML-only scope. Eager owners memoize that fingerprint
by immutable OPC graph allocation; source owners memoize it by source version,
and deferred outputs retain the existing one-time fingerprint, so repeated
composition checks do not re-emit the archive.

Validation performed for this slice:

* `cargo check -p litchi-docx --lib`;
* `cargo check -p litchi-docx --tests` (rerun after the shared cargo lock is
  available);
* `cargo doc -p litchi-docx --no-deps` (rerun with the isolated DOCX doc target);
* `document_limits::raw_xml_policy_admits_selected_depth_nodes_and_attributes_before_layout`;
* `document_limits::raw_output_replay_cannot_widen_limits_and_failed_mutation_is_atomic`;
* `package::tests::graph::package_root_three_way_transfer_history_and_durable_reopen_are_coupled`;
  its changed-untouched-part case has the same main XML and returns
  `document::Error::PatchConflict` (the facade maps this to `invalid_input`),
  while an independently reopened equal-content artifact commits the same
  result;
* `source_backed::composition_binds_reopened_equal_content_untouched_parts_and_source_freshness`;
  it covers the source owner's deferred publication, equal-content reopen,
  changed untouched member refusal, and source revision fence;
* the remaining DOCX composition, managed-owner, and source-backed tests;
* `tools/check_api_glossary.py` and the workspace API debt inventory.

The validation log is kept outside the repository in the coordinator scratch
area; no generated executable or large performance artifact is committed.
