# 0359s — legacy raw replay classification

The legacy raw owners now expose the same replay vocabulary at their public
boundaries.  A patch base mismatch keeps the owning crate's typed error:

| Owner | Base mismatch |
| --- | --- |
| OLE property sets and OLE1 objects | `litchi_cfb::OleError::InvalidFormat` |
| DataSpaces | `litchi_crypto::spaces::Error::Invalid` |
| DOC `RgDofr` and `Selsf` records | `litchi_doc::package::Error::Corrupted` |
| DOC whole-package protection patch | `litchi_doc::package::Error::InvalidFormat` |

These are malformed-input errors at the unified facade.  The shared object
patch error remains `PatchConflict` and classifies as `invalid_input`; a
positional `SourceChanged` remains `io_error`.  The public DOC record APIs use
`PatchConflict` for their semantic patch errors, and the subdocument table
patch uses the same name in place of `SourceConflict`.  Private package and
body-edit conflict names were left unchanged.

The raw replay tests cover an independently reopened source with equal
content and a different source with the typed error above.  The facade tests
also verify the malformed, patch-conflict, and source-changed classifications.

## Validation

On 2026-10-22 these commands ran from the sparse validation tree at
`/home/zhuhe/code/litchi-worktrees/0362rd-substrate-scope`, using the borrowed
`/home/zhuhe/code/litchi-worktrees/uo2-doc-toggle/target` target.  The target
was 4.1 GiB after the run.  Every cargo command used `--offline --locked`,
`CARGO_BUILD_JOBS=1`, and `RUST_TEST_THREADS=1`; logs remain outside the
checkout under `/home/zhuhe/code/litchi-worktrees/.coord/legacy-replay-errors-validation`.

All final validation commands passed:

| Command | Log (bytes, SHA-256) |
| --- | --- |
| `cargo test --offline --locked -p litchi-ole-common --lib property_set::summary_information::tests::filetime_conversions_and_reversible_patches_are_source_checked -- --exact` | `01-ole-summary.log` (1,109; `e8111aa2a105421f0bd9716b01bf89176138f213e2e789be2bb371745bef5cd4`) |
| `cargo test --offline --locked -p litchi-ole-common --lib property_set::document_summary::tests::patch_requires_the_exact_source_and_failed_edits_are_transactional -- --exact` | `02-ole-document-summary.log` (431; `7d62ff078f57a6d036fc01b5340d5cd35905fa9501ddc20f8cc79b81433cc728`) |
| `cargo test --offline --locked -p litchi-ole-common --test ole1 limits_are_checked_before_admission_and_stale_patch_is_rejected -- --exact` | `03-ole1.log` (543; `7c174291db4527b4f3a9c99a4fd7360d6b6cd37ef85b2b79a26bce8eca329e83`) |
| `cargo test --offline --locked -p litchi-crypto --lib spaces::tests::transaction_changes_definition_and_rejects_stale_patch_sources -- --exact` | `04-crypto-spaces.log` (989; `32b575226f66e4749c062e59f17a76d90ddc7059765906708600cb01eebe6b84`) |
| `cargo test --offline --locked -p litchi-doc --lib internal_tests::dofr_edits::changed_dofr_patch_requires_the_complete_source_owner -- --exact` | `05-doc-dofr-foreign.log` (453; `3774e4b94127f6c691b6bfef503066dfd89e6410ddf8d5d8c5a234c2c3fb54d7`) |
| `cargo test --offline --locked -p litchi-doc --lib internal_tests::dofr_edits::reopened_identical_source_with_different_limits_accepts_dofr_patch -- --exact` | `06-doc-dofr-equal.log` (523; `946ffc1265ccd3a9a421f705a90c5bd26e1c514b44b624cd5d683d999149ac07`) |
| `cargo test --offline --locked -p litchi-doc --lib internal_tests::saved_selection::body_saved_selection_edit_is_source_bound_reversible_and_atomic -- --exact` | `07-doc-saved-selection.log` (468; `517e7cbf32e7f4d6bf457057f76356c7aca91cf698d56b4e566fe308e8515870`) |
| `cargo test --offline --locked -p litchi-doc --lib parts::annotation_bookmarks::tests::transaction_is_atomic_and_reversible -- --exact` | `08-doc-annotation-bookmarks.log` (444; `c1f3cbb0220c6a6cf2e1b2a1cddf89a442be2738fea4793d131f6cd989130460`) |
| `cargo test --offline --locked -p litchi-ops --no-default-features --features doc --lib error::tests::legacy_raw_replay_categories_reach_the_facade_unchanged -- --exact` | `09-ops-classifier.log` (497; `961b70d8f39afcfc5e7c38a004438c2092d44467daa7c4dcf72410c5d7957f67`) |
| `cargo check --offline --locked -p litchi-doc -p litchi-ole-common -p litchi-crypto -p litchi-ops --features litchi-ops/doc --all-targets` | `10-all-consumer-check.log` (447; `d32f4ef1f4c0a4760258de51c6de55540daa5deca87e2abf44a30cf56c5b80ed`) |

Two setup attempts failed before the final run and were corrected without
changing the result: the first OLE command lacked the sparse tree's materialized
`Cargo.lock` (`98cadf5d76374f8350a710f62ef8be21ce0e1e5a38a6e66723f3c8651319ac4b`,
262 bytes), and the first DOC command addressed an internal test as a nonexistent
integration target (`24a657b09d0cc776073c5d578aa2916f5fd832e6c18e739e9c62218cadec398e`,
1,215 bytes).  Materializing `Cargo.lock` and using the crate's
`internal_tests::*` target fixed both setup errors.
