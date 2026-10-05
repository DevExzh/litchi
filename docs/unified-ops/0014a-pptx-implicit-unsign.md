# 0014a — litchi-pptx: no implicit `unsign()` (F1 of 0013)

- Status: implemented on `uo/0014a-pptx-unsign`, based on `uo/m0a` at
  `0b79b2f61a` (record 0013, S6b, merged). Follow-up F1 of record
  [0013](0013-signature-policy.md) §6, for `crates/litchi-pptx` only; the
  DOCX and XLSX parts of F1 are separate records.
- Date: 2026-10-05
- Scope: `crates/litchi-pptx` (library, its unit tests, two integration test
  files). No other crate changed; no public signature changed.
- Governing records: ADR 0006 ("Signed input is `Signed<T>`. Editing requires
  consuming it through explicit signature stripping or edit-and-resign"), ADR
  0003 (exact no-ops stay exact), record 0013 §2.1 (the XLSB classification
  this one follows) and §2.2 (the pptx families), UNIFIED_OPS_GOAL
  constraint 1 (typed refusals, exact no-ops).
- `performance_claim: none`. Nothing was timed.

Path prefix: `P/` = `crates/litchi-pptx/src/`. "Base" line numbers are at
`0b79b2f61a`; every other line number is at this record's commit.

## 1. The policy as applied

A changed edit invalidates every package signature, and only the caller may
remove one: `Package::unsign` (`P/package/model.rs:2104`, added by S6b),
`OpcPackage::unsign` on a raw package, or the existing explicit
`PlaceholderSignaturePolicy::Invalidate`. Every package-level edit family of
litchi-pptx now checks, in this order:

1. **exact no-op**: return without touching the package, so a signed package
   republishes its source byte for byte;
2. **refusal**: `refuse_signed_change` (`P/package/signature.rs:36`) returns
   `Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)`, the error S6b's
   opened-presentation commit returns, before the first mutation of the
   caller's package;
3. **mutation**, then **settle**: `settle_unsigned_change`
   (`P/package/signature.rs:57`) replaces each former bare `unsign()` at the
   same position.

"Signed" is `requires_disposition` (`P/package/signature.rs:26`):
`is_signed() || requires_signature_edit_policy()`, so a signature graph that
generic part removal took apart still counts (ADR 0006: generic removal is
not an explicit removal). It is the condition the refusals that predate this
record test: S6b's commit refusal (`P/opened/transaction.rs`) and the
collaboration, read-only-recommended, content-part, InkAction and placeholder
families.

Why the settle step keeps `OpcPackage::unsign` instead of deleting the call:
on a package that step 2 proved unsigned it removes nothing, but it marks the
absent signature graph as explicitly handled and drops the relationship source
captures (`Relationships::retain`), which the families have always done.
Keeping it keeps unsigned packages on the code path they had. Unlike the bare
call, `settle_unsigned_change` refuses a package that still needs a
disposition, so it cannot strip even on a path that missed step 2. A
probe compared the published bytes of unsigned packages before and after
this change: identical (§3, item 8).

## 2. Classification

Classes, as the task defines them:

- **(a)** explicitly authorized, or after a refusal, normalizing only: kept,
  with a comment;
- **(b)** strips silently on a changed edit: refused (§1 step 2), the call
  moved behind the refusal (§1 step 3);
- **(c)** also runs on an exact no-op, where it stripped a package the
  operation did not change: the no-op now returns before the mutation
  (byte comparison of the payload the site would publish with the stored
  one), and a changed edit is refused as in (b).

Counts: 63 sites (62 `unsign()` calls in the families plus the one in
`Package::edit_opc`; `Package::unsign`'s own call is the authorization, not a
site). **(a) 2, (b) 53, (c) 8.** Every (c) site also strips on a changed
edit; one (b) site is test-only (#44).

| # | Base site | Function (reached from) | Class | Action (this commit) |
|---|---|---|---|---|
| 1 | `P/comments/package.rs:314` | `add_presentation_comment_author`, new graph | b | refuse `:321`, settle `:323` |
| 2 | `P/comments/package.rs:518` | `add_presentation_comment` | b | refuse `:538` / `:547`, settle `:564` |
| 3 | `P/comments/package.rs:611` | `remove_presentation_comment`, last comment of a slide | b | refuse `:671`, settle `:679` |
| 4 | `P/comments/package.rs:672` | `commit_legacy_authors` (add / update / replace / remove / reorder author, add comment) | c | no-op `:748`, refuse `:751`, settle `:753` |
| 5 | `P/comments/package.rs:686` | `commit_legacy_slide_comments` (update / replace / reorder comments, remove a non-last comment) | c | no-op `:768`, refuse `:771`, settle `:773` |
| 6 | `P/model3d/package.rs:116` | `put` (`Package::put_model3d`) | b | refuse `:78`, settle `:118` |
| 7 | `P/model3d/package.rs:141` | `remove` (`Package::remove_model3d`) | b | refuse `:139`, settle `:144` |
| 8 | `P/table/style/validation.rs:110` | `put`, replacement (`Package::put_styles`, `Transaction::put_table_styles`) | b | refuse `:116`, settle `:119` |
| 9 | `P/table/style/validation.rs:132` | `put`, new catalog | b | refuse `:139`, settle `:142` |
| 10 | `P/table/style/validation.rs:168` | `remove` (`Package::remove_styles`, `Transaction::remove_table_styles`) | b | refuse `:183`, settle `:186` |
| 11 | `P/chart/package.rs:91` | `chart::add` (`Transaction::add_chart`) | b | refuse `:91`, settle `:100` |
| 12 | `P/tag/package/codec.rs:235` | `tag::put`, replacement (`Package::put_tags`) | b | refuse `:216`, settle `:244` |
| 13 | `P/tag/package/codec.rs:271` | `tag::put`, new attachment | b | refuse `:265`, settle `:281` |
| 14 | `P/tag/package/codec.rs:341` | `tag::remove` (`Package::remove_tags`) | b | refuse `:336`, settle `:359` |
| 15 | `P/tag/shape/codec.rs:185` | `tag::shape::put`, replacement (`Package::put_shape_tags`) | b | refuse `:162`, settle `:194` |
| 16 | `P/tag/shape/codec.rs:220` | `tag::shape::put`, new attachment | b | refuse `:210`, settle `:230` |
| 17 | `P/tag/shape/codec.rs:294` | `tag::shape::remove` (`Package::remove_shape_tags`) | b | refuse `:289`, settle `:312` |
| 18 | `P/master_layout/package.rs:935` | `invalidate_signatures`, called by `add_slide_master` (base `:155`), `add_slide_layout` (`:250`), `remove_slide_layout` (`:368`) | b | refuse `:125` / `:232` / `:390`, settle `:164` / `:267` / `:400`; the helper is removed |
| 19 | `P/master_layout/placeholder.rs:419` | `apply_patch_with_policy`, `SignaturePolicy::Invalidate` | a | kept `:422`: the caller chose Invalidate; `Reject` refuses before it; a no-op returns first |
| 20 | `P/master_layout/placeholder.rs:985` | `apply_slot_patch_with_policy`, `SignaturePolicy::Invalidate` | a | kept `:990`, as #19 |
| 21 | `P/master_layout/placeholder.rs:1035` | `store_placeholder_shape_source` (`store_placeholder_shape`, `Package::store_placeholder_shape`) | b | refuse `:1045`, settle `:1049`; the "always invalidated" comment is replaced |
| 22 | `P/presentation_properties/package.rs:131` | `put_math_to_package` (`Package::put_math_properties`) | b | refuse `:102`, settle `:140` |
| 23 | `P/presentation_properties/package.rs:171` | `remove_math_from_package` (`Package::remove_math_properties`) | b | refuse `:167`, settle `:188` |
| 24 | `P/presentation_properties/metadata/changes/package.rs:32` | `changes::store` | b | refuse `:35`, settle `:41` |
| 25 | `…/metadata/changes/package.rs:127` | `changes::remove` | b | refuse `:128`, settle `:158` |
| 26 | `…/metadata/changes/package.rs:156` | `replace_snapshot` (`changes::Patch::apply`, `apply_patch`, `apply_commit`) | b | refuse `:176`, settle `:188` |
| 27 | `…/metadata/structure/package.rs:65` | `structure::apply_patch` / `apply_commit` / `Patch::apply` | b | refuse `:63`, settle `:74` |
| 28 | `…/metadata/structure/codec.rs:68` | `structure::store`, and through it every custom-show and section CRUD function and `synchronize_after_slide_mutation` | b | settle `:78` (before the mutation, where the call was; it refuses) |
| 29 | `…/metadata/events/package.rs:55` | `events::store` | b | refuse `:58`, settle `:64` |
| 30 | `…/metadata/events/package.rs:78` | `events::remove` | b | refuse `:89`, settle `:95` |
| 31 | `…/metadata/events/package.rs:109` | `events::apply_patch` / `apply_commit` / `Patch::apply` | b | refuse `:122`, settle `:134` |
| 32 | `…/metadata/guides/package.rs:95` | `guides::apply_patch` (`store`, `remove`, `apply_commit`, `Patch::apply`) | b | refuse `:107`, settle `:118` |
| 33 | `…/metadata/designer_tags/package.rs:179` | `designer_tags::apply_patch` (`store`, `remove`, `Commit::apply`, `Patch::apply`, `Package::put/remove_slide_designer_tags[_with_limits]`) | b | refuse `:191`, settle `:209` |
| 34 | `…/metadata/custom_show/package.rs:82` | `custom_show::apply_patch` (`store`, `remove`, `apply_commit`, `Patch::apply`) | b | refuse `:87`, settle `:98` |
| 35 | `P/modern_comments/commands.rs:147` | `update_modern_comment_changes`; a callback that changes nothing rewrote and stripped | c | no-op `:156` (returns `Ok(true)`, as before), refuse `:159`, settle `:161` |
| 36 | `P/modern_comments/package.rs:247` | `add_modern_comment`, new per-slide part | b | refuse `:241`, settle `:256` |
| 37 | `P/modern_comments/package.rs:369` | `remove_modern_comment`, last comment of a part | b | refuse `:392`, settle `:400` |
| 38 | `P/modern_comments/package.rs:674` | `validate_and_commit_modern_comment_part` (add to an existing part, update / replace / reorder comments, replies, extensions, remove a non-last comment) | c | no-op `:755`, refuse `:758`, settle `:760` |
| 39 | `P/modern_comments/package.rs:1120` | `add_modern_comment_author`, new author part | b | refuse `:1200`, settle `:1215` |
| 40 | `P/modern_comments/package.rs:1211` | `remove_modern_comment_author`, last author | b | refuse `:1320`, settle `:1328` |
| 41 | `P/modern_comments/package.rs:1267` | `commit_modern_comment_authors` (add to an existing part, update / replace / reorder authors, remove a non-last author) | c | no-op `:1392`, refuse `:1395`, settle `:1397` |
| 42 | `P/notes/package.rs:403` | `notes::apply_patch` / `apply_commit` (`Package::put_notes`, `apply_notes_commit`, `apply_notes_patch`) | b | refuse `:377`, settle `:405` |
| 43 | `P/notes/package.rs:1182` | `apply_notes_removals` (`Package::remove_notes`, `clear_notes`) | b | refuse `:1176`, settle `:1186` |
| 44 | `P/notes/package.rs:1395` | `put_changed`, `#[cfg(test)]` only | b | refuse `:1391`, settle `:1401` |
| 45 | `P/change_tracking/package.rs:50` | `apply_patch` / `apply_commit` (`Package::apply_change_tracking_*`) | b | refuse `:42`, settle `:52` |
| 46 | `P/shape/classification/package.rs:71` | `put` (`Package::put_shape_classification`) | b | no-op moved before the mutable borrow `:61`, refuse `:64`, settle `:75` |
| 47 | `P/shape/classification/package.rs:118` | `remove` (`Package::remove_shape_classification`) | b | no-op `:113`, refuse `:116`, settle `:127` |
| 48 | `P/shape/theme/package.rs:40` | `theme::add` | b | refuse `:43`, settle `:49` |
| 49 | `P/shape/theme/package.rs:69` | `theme::attach` (always rewrites the master's theme relationship) | b | refuse `:74`, settle `:86` |
| 50 | `P/shape/theme/package.rs:114` | `theme::put_colors`; the palette already stored rewrote and stripped | c | no-op `:139`, refuse `:142`, settle `:144` |
| 51 | `P/shape/theme/package.rs:136` | `theme::put_fonts`, as #50 | c | no-op `:174`, refuse `:177`, settle `:179` |
| 52 | `P/shape/theme/package.rs:167` | `theme::put_override`, existing override; the stored override rewrote and stripped | c | no-op `:215` (read-back kept), refuse `:221`, settle `:226` |
| 53 | `P/shape/theme/package.rs:183` | `theme::put_override`, new override | b | refuse `:230`, settle `:243` |
| 54 | `P/shape/theme/package.rs:237` | `theme::remove_override` | b | refuse `:293`, settle `:305` |
| 55 | `P/shape/designer/properties.rs:289` | `apply_properties` (`Package::put/remove_shape_designer_properties[_with_limits]`) | b | source check before the mutable borrow, no-op `:288`, refuse `:291`, settle `:296` |
| 56 | `P/shape/zoom/package.rs:51` | `store` (`Package::put_zooms`, `remove_zooms`) | b | refuse `:50`, settle `:53` |
| 57 | `P/shape/designer/package.rs:70` | `put` (`Package::put_shape_design_element`) | b | no-op `:60`, refuse `:63`, settle `:74` |
| 58 | `P/shape/designer/package.rs:117` | `remove` (`Package::remove_shape_design_element`) | b | no-op `:112`, refuse `:115`, settle `:126` |
| 59 | `P/presentation/embedded/ole/slide/package.rs:82` | `ole::slide::apply_patch` / `apply_commit` / `Patch::apply` | b | refuse `:89`, settle `:91` |
| 60 | `P/presentation/embedded/controls/slide/package.rs:102` | `controls::slide::apply_patch` / `apply_commit` / `Patch::apply` | b | refuse `:109`, settle `:111` |
| 61 | `P/presentation/embedded/ole/authoring.rs:95` | `ole::add` | b | refuse `:81`, settle `:104` |
| 62 | `P/font/package.rs:252` | `put_raw` (`font::put`, `font::remove`, `Package::put_fonts`, `remove_fonts`) | b | refuse `:252`, settle `:254` |
| 63 | `P/package/model.rs:1915` | `Package::edit_opc`: a signed source whose closure changed the published bytes was unsigned | b | refused instead (`:2138`, check `:2153`); see §3 item 4 |

The no-op of every (b) site was already exact, except the shape-span puts and
removes (#46, #47, #55, #57, #58): their no-op returned without `unsign()`
but after `get_part_mut`, which revokes the exact-source authorization, so an
unchanged signed package could no longer be published (`to_bytes` refused
it). The no-op check now precedes the mutable borrow (pinned by
`shape_put_no_ops_keep_a_signed_package_exact`; with the borrow moved back
in front of the check, that test fails with the writer's refusal).

### 2.1 Found while classifying, not changed here

- **Strippers outside litchi-pptx reached through its API.**
  `Package::put_custom_props` / `remove_custom_props` call
  `litchi_ooxml_common::custom::Props::write_for`, which unsigns a changed
  package (`crates/litchi-ooxml-common/src/custom/package.rs:97`, `:121`,
  `:126`); `presentation::embedded::vba::{store, remove}` call
  `store_project_graph` / `remove_project_graph`
  (`crates/litchi-ooxml-common/src/vba.rs:211`, `:267`). The same crate has
  further sites (ribbon, custom XML, core properties). They are shared with
  DOCX and XLSX and belong to a litchi-ooxml-common record (F1-common, §6).
- **Late refusals.** Some pptx mutators change a signed package in place
  without `unsign()` and without a refusal, so the package refuses only when
  it is saved: `store_presentation_comments`, `store_modern_comment`,
  `store_modern_comment_authors`, `media_parts::store`
  (`Transaction::store_media`) and `slide_sync::store`
  (`Package::store_slide_sync`). No signature is stripped; the refusal comes
  late, as record 0013 F2 records for XLSB task panes. Not exhaustive.
- **Shape-span codecs need self-contained namespaces.** The classification
  and design-element codecs read one shape's XML on its own, so a shape whose
  `p:` prefix is declared on the slide root (both producer-signed fixtures)
  fails with "classification shape root is not PresentationML"; the
  modern-comment command parser needs `xmlns:pc2` on `pc2:cmMkLst` itself.
  Unrelated to signatures; the tests use self-contained XML.

## 3. Behaviour changes

1. **Refusal.** Every (b) and (c) site refuses a changed edit of a signed
   package with `Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)`
   before the caller's package is mutated; before, the edit succeeded and
   silently removed every signature. This covers the public functions on
   `OpcPackage` of the comment, modern-comment, tag, table-style, chart,
   master/layout/placeholder, math, Changes Information, presentation
   structure (custom shows, sections), slide-show event, guide, Designer-tag,
   custom-show, theme, OLE, ActiveX and font families, their `Patch` /
   `Commit` `apply` methods, and the `Package` methods built on them (notes,
   fonts, tags, shape tags, slide Designer tags, math, zooms, shape
   classification, design element, Designer properties, model3d, table
   styles, masters and layouts, placeholder shapes, change tracking). Call
   `Package::unsign` first (or `OpcPackage::unsign` on a raw package). Base
   evidence: of the 44 changed edits of the family table on the two
   producer-signed fixtures, 40 stripped and 4 were already refused
   (`results/0014a/base-signed-families.txt`).
2. **Exact no-ops.** The (c) sites publish nothing when the payload they
   would write equals the stored bytes: legacy-comment author and comment
   updates and reorders that change nothing, the same for modern comments,
   authors, replies and extensions, a command callback that changes nothing,
   a theme palette, font set or override that is already stored. A signed
   package stays byte-identical; before, it was stripped. A semantic no-op
   whose re-serialization differs from the stored bytes (for example a
   producer-written author list put back unchanged) is a change and is
   refused on a signed package.
3. **Shape-span no-ops** (§2, last paragraph) keep a signed package
   publishable and byte-identical.
4. **`Package::edit_opc`.** A closure that changes a signed package is
   refused and rolled back unless it removes the signature explicitly
   (`OpcPackage::unsign`), or the caller called `Package::unsign` first. A
   closure that changes nothing keeps the signature byte for byte, as
   before. Before: a changed result that could still be serialized and
   differed from the source was unsigned, which removed an in-memory (not
   ingressed) signature graph, as the package-encryption test pinned; a
   closure that called a stripping family had already lost the signature
   inside it. A signed source that the closure changed without unsigning was
   refused by the writer, and still is. The documentation said nothing about
   signatures; it now states this.
5. **Opened-presentation transaction.** The verbs that stage through these
   families (`put/remove_table_styles`, `add_chart`, `add_slide_master`,
   `add/remove_slide_layout`, `add_comment_author`, `add/remove_comment`, and
   the thirteen modern-comment verbs) refuse a signed working package while
   staging; before, they staged a copy without the signature and S6b's
   `commit` refused. Same error, earlier; `Transaction` documents it.
6. **`master_layout::store_placeholder_shape`** (and its facade method), the
   compatibility API whose comment said it "has always invalidated
   signatures", refuses instead.
7. **Documentation.** 162 public functions and methods state the refusal and
   name `Package::unsign` (raw functions also `OpcPackage::unsign`;
   transaction verbs: before the transaction starts); `Package::unsign`,
   `Package::edit_opc`, `Transaction` and `font::put` were rewritten where
   they described signature removal as a side effect.
8. **Unsigned packages: unchanged.** The refusal never fires for them, and
   the settle step runs where `unsign()` ran. A probe ran every no-op and
   changed edit of the family table, plus the legacy-comment class c no-ops,
   on unsigned packages with the base and the new sources: the 94 published
   outputs are byte-identical (`results/0014a/unsigned-outputs.txt`). The
   one difference not visible in bytes: a byte-level no-op no longer
   revokes the exact-source authorization or resets the OPC signature-policy
   flags.

## 4. Tests

New, `crates/litchi-pptx/tests/package_signature_policy.rs` (10 tests):

- `every_family_refuses_a_signed_package_until_it_is_unsigned`: a table of
  25 families (legacy comments, modern comments, notes, slide tags, shape
  tags, change tracking, shape classification, design element, Designer
  properties, zoom and model3d owners, themes, math, sections, custom shows,
  slide-show events, extended guides, slide Designer tags, Changes
  Information, OLE objects, fonts, table styles, masters / layouts /
  placeholders, charts, `Package::edit_opc`, the opened-presentation
  transaction), each on both producer-signed fixtures
  (`test-data/poi/test-data/xmldsign/hello-world-signed.pptx`,
  `ms-office-2010-signed.pptx`): the changed edit is refused and the package
  still publishes its source; the exact no-op publishes the source; after
  `Package::unsign` the edit succeeds, the output has no signature part, and
  the reopened package shows the change. Families whose content the fixtures
  lack are prepared after an explicit unsign and signed again in-test (an
  origin part, one XML signature and their relationships, as record 0013's
  XLSB tests do); notes, zoom, model3d and guides have no change row there
  and are covered below or by their no-op row.
- `legacy_comment_no_ops_keep_a_signed_package_exact`,
  `modern_comment_no_ops_keep_a_signed_package_exact`,
  `modern_comment_command_no_op_keeps_a_signed_package_exact`,
  `theme_no_ops_keep_a_signed_package_exact`: the (c) sites; each fails on
  the base sources at its byte-identity assertion.
- `shape_put_no_ops_keep_a_signed_package_exact`: §2's mutable-borrow fix.
- `notes_refuse_a_signed_package_until_it_is_unsigned` (POI `prProps.pptx`,
  signed in-test), `model3d_refuses_a_signed_package_until_it_is_unsigned`.
- `raw_edit_that_unsigns_explicitly_publishes_without_a_signature`,
  `generic_removal_of_signature_parts_does_not_authorize_a_change` (removing
  the signature parts and relationship by hand leaves the OPC policy
  requirement; the tag family still refuses until `OpcPackage::unsign`).

New unit test: `package::signature::tests::settling_refuses_a_signed_package_and_never_removes_its_signature`.

Rewritten, because they pinned the silent strip; each now asserts the
refusal, that the package is unchanged and still signed, and the edit after
an explicit unsign:

| Test | Was | Now |
|---|---|---|
| `font::tests::present_empty_and_loaded_false_flag_round_trip_exactly` | `remove` of a signed package succeeded | refused, then succeeds after `unsign` |
| `font::tests::generated_crud_allocates_collisions_and_preserves_unknown_xml_atomically` | changed `put` returned `true` and `!is_signed()` | refused (parts and XML unchanged), then succeeds |
| `change_tracking::tests::facade_publishes_and_reverses_exact_source_checked_commit` | changed commit on a signed package published | refused (state unchanged), `Package::unsign`, published |
| `notes::tests::consuming_put_moves_changed_xml_and_preserves_signed_no_ops` | changed `put` stripped | renamed `…_refuses_signed_changes_until_unsigned`: refused, graph unchanged, then succeeds |
| `presentation_properties::metadata::designer_tags::tests::inserts_replaces_and_removes_without_touching_opaque_bytes` | `remove` stripped | refused (owner bytes unchanged), then succeeds |
| `table::style::tests::package_crud_round_trips_all_main_profiles_and_both_conformances` | changed `put` stripped | refused, then succeeds |
| `table::style::tests::semantic_noop_keeps_preserved_and_opaque_whitespace_distinct` | changed opaque whitespace stripped | refused (source unchanged), then succeeds |
| `tag::shape::tests::shape_crud_is_atomic_noop_safe_and_move_based` | `remove` stripped | refused (owner unchanged), then succeeds |
| `tag::tests::anchored_crud_is_strict_profiled_and_signature_safe` | changed `put` and `remove` stripped | both refused (owner and tag part unchanged), then succeed |
| `tag::tests::direct_owner_mutates_active_mce_fallback_and_preserves_inactive_source` | changed `put` and `remove` stripped | both refused, then succeed |
| `tests/package_encryption.rs::raw_opc_noop_preserves_signature_but_changed_edit_unsigns` | changed `edit_opc` stripped | renamed `…_and_changed_edit_requires_explicit_unsign`: refused and rolled back; succeeds when the closure calls `OpcPackage::unsign` |

S6b's `tests/opened_signature_policy.rs` and
`opened::tests::changed_commit_refuses_a_signed_package_until_it_is_unsigned`
pass unchanged.

## 5. Gates

On this branch, with `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0` and `3rdparty`
symlinked from the main checkout for the test runs (removed afterwards):

- `cargo fmt --check -p litchi-pptx`: clean.
- `cargo clippy -p litchi-pptx --all-targets --all-features -- -D warnings`:
  clean.
- `cargo test -p litchi-pptx --all-features --no-fail-fast`: 1,325 passed,
  3 ignored, 0 failed (base: 1,314 passed, 3 ignored).
- `RUSTDOCFLAGS=-Dwarnings cargo doc -p litchi-pptx --no-deps --all-features`:
  clean.
- `cargo test -p litchi-ops --all-features signature`: 26 passed, including
  `tests::signature::every_signed_input_follows_the_explicit_signature_policy`
  and `an_unlocked_signed_package_is_not_unsigned_into_the_clear`.
- `cargo check -p litchi-ops -p litchi --all-features`: clean.

## 6. Follow-ups

- **F1-common**: classify and refuse the litchi-ooxml-common strippers
  (custom properties, VBA, ribbon, custom XML, core properties; §2.1), which
  litchi-pptx reaches through `Package::put_custom_props`,
  `remove_custom_props` and `presentation::embedded::vba::{store, remove}`
  and which DOCX and XLSX share.
- **F1-late**: refuse up front in the late-refusal mutators of §2.1 (the pptx
  counterpart of record 0013 F2).
- **F1-semantic**: let a semantic no-op whose serialization differs from the
  stored bytes publish nothing (§3 item 2), where the family's model is
  lossless enough to prove it.
- The private refusal helpers that predate `P/package/signature.rs`
  (collaboration, read-only-recommended, content parts, InkAction,
  placeholders, opened transaction) can call it; behaviour is the same.

## 7. Not claimed

- Nothing was timed. The refusal adds an `is_signed` scan of the package's
  relationships and parts per changed edit; not measured.
- No signature is verified; signatures are detected and removed on request.
- The probes of §3 were temporary tests, deleted after the runs; their
  outputs and method are in `results/0014a/`. Their conclusions are pinned by
  the tests of §4, except item 8 (unsigned identity), which rests on the
  probe and on the unchanged unsigned tests of the suite.
- §2.1 is read from code; its late-refusal list is not exhaustive.
