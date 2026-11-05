# 0362ul87aj — remove stale OOXML publishers

This source closure starts from `6e6b0e5bfab50dd96668f26f197b530df3c99630`
(tree `6793f75c0ac86c42a42ae9fdbd53234dd3992242`). It completes the 87aj
review by removing the old crate-private publishers and migrating every
remaining consumer in the same change.

The DOCX bibliography fixtures now create their Custom XML occurrence through
`custom_xml::Edit::insert` and publish with its `Commit`, preserving the
canonical source-checked graph path. The single core-properties refusal test
now stages a complete `properties::Edit::set_metadata` and asserts the typed
`Field::ContentType` refusal without mutating its source package.

The obsolete `custom::Props::write`/`write_for`,
`custom_xml::add` and its rollback/name helpers,
`properties::write`, and the extended-properties package `write`, `clear`,
and relationship-ID publisher are gone. Their former semantic publication
responsibility remains in the shared Edit/Commit paths; low-level XML and
graph codecs remain available for their test fidelity surfaces. The
`properties::write_core` re-export is now test-only because no release
consumer remains after the old wrapper was removed. Stale documentation links
were updated and no aliases or deprecated twins were added.

The preserved F5 receipts remain unchanged:

* `e13bd7e843` stopped at scoped formatting; log SHA-256
  `394fda5b21fc3c4b8c767522da24b59fc955c96a5249da3b3bb5fbbf0a914a4f`.
* `a3eff9e1a6` stopped at `check-common` with the stale `clear` call; log
  SHA-256 `029639ec02ad68ec8c7851390781263fe3a2138a40dc0b447d616ee503aec5c0`.
* `27be5389f7` stopped at `check-common` with the 22 release dead-code
  diagnostics; log SHA-256
  `e1ee9fdb52e67b0a4ba8cce71834b1de692eaaa6d0bbf6e0d220ff8b45ef8c4b`.

`git diff --check`, the exact scoped rustfmt command over all changed Rust
paths, and stale-symbol audits pass. A preliminary API-glossary invocation
was not an acceptance run: it invoked Cargo/rustdoc and stopped at the
release-only `write_core` unused-import diagnostic; the test-only re-export
repair above resolved that source issue. No F5, Cargo gate, rustdoc gate,
test, native, or performance run is claimed for this tip; the next attempt
must bind this commit and refreshed fixture identities in a fresh preflight.
