# 0359s — legacy raw replay classification

The legacy raw owners expose the same replay vocabulary at their public
boundaries. A patch base mismatch keeps the owning crate's typed error:

| Owner | Base mismatch |
| --- | --- |
| OLE property sets and OLE1 objects | `litchi_ole_common::Error::PatchConflict` |
| DataSpaces | `litchi_crypto::spaces::Error::PatchConflict` |
| DOC `RgDofr` and `Selsf` records | `litchi_doc::package::Error::PatchConflict` |
| DOC whole-package protection patch | `litchi_doc::package::Error::PatchConflict` |

Malformed source or replacement data keeps its existing structural error:
`litchi_cfb::OleError::InvalidFormat`, `litchi_crypto::spaces::Error::Invalid`,
or `litchi_doc::package::Error::Corrupted`, as appropriate. The shared object
patch error remains `PatchConflict` and classifies as `invalid_input`; a
positional `SourceChanged` remains `io_error`. The public DOC record APIs use
`PatchConflict` for their semantic patch errors, and the subdocument table
patch uses the same name in place of `SourceConflict`. Private package and
lower-level transaction conflicts remain private; the public body-text owner
uses `PatchConflict` for stale replay and reserves `Conflict` for three-way
semantic overlap.

The original 0359s audit had incorrectly retained malformed variants for the
seven raw replay boundaries. The 0359sb correction changes those boundaries to
typed conflicts, migrates their consumers, and keeps malformed-input tests
separate.

## Validation

On 2026-10-22 these commands ran from the sparse validation tree at
`/home/zhuhe/code/litchi-worktrees/0362rd-substrate-scope`, using the borrowed
`/home/zhuhe/code/litchi-worktrees/uo2-doc-toggle/target` target. Every cargo
command used `--offline --locked`, `CARGO_BUILD_JOBS=1`,
`RUST_TEST_THREADS=1`, and a bounded `TMPDIR` under
`(deleted-coord-artifact)`.
Logs remain outside the checkout in that directory.

All final validation commands passed:

| Command | Log (bytes, SHA-256) |
| --- | --- |
| `cargo test --offline --locked -p litchi-ole-common --lib property_set::summary_information::tests::filetime_conversions_and_reversible_patches_are_source_checked -- --exact` | `11-ole-summary-patchconflict.log` (2,163; `e282c431bea94ebd0e0859defa393b298059f967e3e08f705dc9240ad15a787c`) |
| `cargo test --offline --locked -p litchi-ole-common --lib property_set::document_summary::tests::patch_requires_the_exact_source_and_failed_edits_are_transactional -- --exact` | `12-ole-document-summary-patchconflict.log` (443; `ce790c5bd5624bdf509d35b9cc51c008e434580dfe124926fd42407f40ba4a01`) |
| `cargo test --offline --locked -p litchi-ole-common --test ole1 limits_are_checked_before_admission_and_stale_patch_is_rejected -- --exact` | `13-ole1-patchconflict.log` (502; `a915b72298dfc025713d4c6da908f454abe294dddeb5ec4a024ab346c2879fcb`) |
| `cargo test --offline --locked -p litchi-crypto --lib spaces::tests::transaction_changes_definition_and_rejects_stale_patch_sources -- --exact` | `14-crypto-spaces-patchconflict.log` (1,584; `345a5ae9498c1633cee09251af0e6cb4f8b39f4d08b5e5cdd974952a96b14980`) |
| `cargo test --offline --locked -p litchi-doc --lib internal_tests::dofr_edits::changed_dofr_patch_requires_the_complete_source_owner -- --exact` | `15-doc-dofr-patchconflict.log` (2,680; `d9d1a3613774985d08faad5bee06d9fc6421a0400bc6aa1627051219a0b1ee64`) |
| `cargo test --offline --locked -p litchi-doc --lib internal_tests::dofr_edits::reopened_identical_source_with_different_limits_accepts_dofr_patch -- --exact` | `16-doc-dofr-equal-source.log` (426; `f6f9191d03492c9663bf2e3170d6d87da8481ba9849aefcf9bbabf1dd39f7d10`) |
| `cargo test --offline --locked -p litchi-doc --lib internal_tests::saved_selection::body_saved_selection_edit_is_source_bound_reversible_and_atomic -- --exact` | `17-doc-saved-selection-patchconflict.log` (428; `996e6c6439355dfc1be8ae3951ee85c41e52fa35c41c167aae562732c2ce8b17`) |
| `cargo test --offline --locked -p litchi-doc --lib parts::annotation_bookmarks::tests::transaction_is_atomic_and_reversible -- --exact` | `18-doc-annotation-package-patchconflict.log` (404; `dc7cead50d2f28167664ede2dbc45748babcd318cf8a68391a65bcfccaaae412`) |
| `cargo test --offline --locked -p litchi-ops --no-default-features --features doc --lib error::tests::legacy_raw_replay_categories_reach_the_facade_unchanged -- --exact` | `19-ops-classifier-patchconflict.log` (2,509; `c4541339ede3182fd4fbb77654646a53b262bdb2903aac604e19f65fbfceacea`) |
| `cargo check --offline --locked -p litchi-doc -p litchi-ole-common -p litchi-crypto -p litchi-ops --features litchi-ops/doc --all-targets` | `20-all-consumers-check.log` (72; `f358a10e8c6e80056a790635897e0eb8d390ac48a051961384a880aef2ae9be9`) |

One initial all-consumer check used the default `/tmp` and failed before
compilation because its tmpfs had no free inodes. The retry used the bounded
scratch directory and passed. The retained failure log is
`failure-03-tmp-inode.log` (473 bytes,
`121ae808a82d62dff19a2a7943fe9a6da7969ac209ce7a62dea8121bb828b51e`).

The post-change error-layout probe printed
`litchi_ole_common::Error=48 litchi_crypto::spaces::Error=40
litchi_doc::Error=48`; its log is `22-error-sizes.log` (367 bytes,
`4734854bd949e48f9b3aa9b5dd7d4094053ae4762d9a3497010bc9181cde5d90`).
