# 0357 — W3: PPT on the record 0351 contract

- Status: implemented on `uo2/w3-ppt`, cut from `feat/unified-ops-api` at
  `e652864de8`, merged with trunk `5f6fb75164`.
- Date: 2026-10-20.
- Input: owner decision [0350](0350-owner-api-debt-decision.md) items 16 and
  18; contract [0351](0351-api-contract.md) §2, §3.5, §4.3, §5 (wave W3).
- Owned: `crates/litchi-ppt/**`; the PPT hunks of
  `crates/litchi-ops/src/{error.rs, error/tests.rs, format/ppt.rs,
  format/signature.rs, open/cfb.rs, raw/validate.rs, tests/signature.rs}` and
  two ops PPT tests; `tools/perf-baseline/src/{lib.rs, unified_ops.rs}`;
  `tools/native-resave/src/main.rs`; the `ppt-*` profiles of
  `tools/api_glossary.json`; the W3 line of `docs/API_GLOSSARY.md` §2a;
  `results/0289/crud-checklist.tsv` (re-blessed, §6).
- `performance_claim`: measured (§7).

## 1. Summary

`litchi-ppt` now conforms to record 0351 §2:

- every public scoped edit is a package-level family
  `x::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` reached through
  `Package::edit_x(&self, ..)`, whose commit publishes a new `Package` with
  `Diagnostics`, replayed by `Package::apply_x(&self, &x::Patch) ->
  x::Commit`;
- no public `Transaction`, `rollback`, `is_changed`, `into_snapshot` or
  `*Editor` name remains; the gate bans them crate-wide;
- no public `*_with_limits` name remains; limits travel in `OpenOptions` and
  the owners keep them;
- the options getters are `<noun>_value()`;
- typed refusals expose `kind()`; scoped staging refusals are
  content-free `Refused(Refusal)` values;
- the chart draft lives in `builder`.

Every PPT round-trip, CRUD, corpus and no-op test passes with byte-identical
outputs. Consumers were migrated in the same commits.

## 2. Scoped families

All families share the crate-private plumbing in
`crates/litchi-ppt/src/scoped.rs`:

- `Base::capture` takes the exact package bytes and refuses an encrypted
  package (`Error::Package(EncryptionPolicy)`), as the primary edit does;
- whole-artifact families publish candidate bytes (`Base::publish`); record
  families publish one persisted owner record through the existing
  append-only publisher (`slide_order::records::publish`, now with an
  `Owner::Persist(id)` variant for masters), which proves every other record
  keeps its bytes and the slide list is unchanged;
- an exact no-op commit returns the source package (byte-identical,
  signatures kept); a changed commit of a signed package is `Error::Signed`;
- a family patch is exact and record-sized since
  [0357a](0357a-ppt-record-sized-patches.md) (endpoint length and SHA-256
  plus the changed 512-byte ranges; it was whole-artifact source and target
  bytes), `inverse` swaps the endpoints, and `apply_x` refuses any other
  package with `Error::PatchConflict`;
- `family!("noun")` declares `Diagnostics` (`litchi_core::edit`), `Commit =
  litchi_core::edit::Commit<Package, Patch>`, `LimitKind` (no caller-selected
  dimension), `Error {Package, Refused, PatchConflict, Signed}`, `Result` and
  `Patch {is_noop, inverse, source_len, target_len, changed_bytes}`
  (`before_bytes`/`after_bytes` until 0357a).

| Family | Entry | Was |
|---|---|---|
| `font` | `edit_fonts()` | `FontSnapshot`/`FontTransaction`/`FontCommit`/`FontPatch` |
| `broadcast` | `edit_broadcast(Position)` | `broadcast::{Snapshot, Transaction, Commit, Patch}` on one container |
| `picture_bullets` | `edit_picture_bullets()` | `PictureBulletTransaction`, `picture_bullets::Transaction` |
| `animation` | `edit_animations()` | `animation::Editor` (`finish() -> Vec<u8>`, `is_changed`) |
| `document_comparison` | `edit_document_comparison()` | record `Editor` and package `Editor`/`PackageCommit` |
| `external_media` | `edit_external_media()` | `external_media::Transaction` |
| `diagram` | `edit_diagrams(slide)` | `SlideEditor`/`SlideCommit`/`SlidePatch` and the BuildList `Transaction` |
| `slide_sync` | `edit_slide_sync(slide)` | `SlideSyncEditor`/`SlideSyncCommit` |
| `client_data` | `edit_client_data(Target)` | `client_data::Transaction` |
| `master::metadata` | `edit_master_metadata(Kind, Position)` | `master::metadata::Editor` (+ `template::Editor`) |
| `master_layout` | `edit_master_layout(Kind, Position)` | `master_layout::Transaction` |
| `embedded::object` | `edit_objects()` | `litchi_ppt::Editor` (`embedded::object::Editor`, `finish() -> Vec<u8>`) |

Each `apply_x` mirrors its `edit_x`. The record-level staging owners stay
crate-private (renamed `RecordEdit`/`RecordCommit`/`RecordPatch` where the
family names would clash), lose `rollback`, `finish`, `is_changed`,
`snapshot` and alias setters, and keep their unit tests. Their `Snapshot`
types stay as public read views.

Duplicates of the primary edit became crate-private instead of families
(record 0351 §3.5 left the choice per family):

- `slide_order::{Transaction, Commit, Patch, Change*, Composition, History,
  PreparedMove, ThreeWayPlan}` and `text_edit::{Snapshot, Transaction,
  Commit, Patch, SourceTransaction, SourceBacked*}` — `Package::edit` is the
  one public edit route; `slide_order::Snapshot` (whole-artifact read view)
  and `text_edit::SourceSnapshot` (source-backed text reader) stay public,
  with the vocabulary the primary verbs take and the typed errors;
- `client_anchor` (`set_shape_anchor`), `document_structure` (slide moves,
  inserts, removals), `package::property_set` (behind `package::metadata`).

[0357a §7](0357a-ppt-record-sized-patches.md#7-follow-up-restoring-the-routes-record-0357-made-crate-private)
restores public routes for the durable, history, composition and three-way
surfaces (on `package::edit`), the source-backed text splice (the
`text_edit` family) and the master and custom-table-style edits (primary
`Edit` verbs).

Drafts: `chart::PackageEditor` became the value draft `chart::GraphPackage`
(`chart::Snapshot` → `chart::View`, `edit_package` → `draft_package`); a chart
replacement stages into `embedded::object::Edit` and publishes a `Package`.
`embedded::storage::Editor` became `storage::Draft`. The writer chart draft is
`builder::{Chart, ChartKind, ChartSeries}` (§2.1 rule 7).

Verified for §3.5: no writer depended on `rollback`; the only non-test callers
were the two `master::metadata` delegations, deleted with the editors.

## 3. Primary edit and owners

- `Package::apply(&Patch) -> package::edit::Commit` and
  `Package::apply_metadata(&Patch) -> metadata::Commit` (were `-> Package`).
- `EditOptions::get_max_operations` → `max_operations_value(&self)`;
  `OpenOptions::{get_source_limits, get_password, get_leniency}` →
  `*_value()`.
- `package::edit::Edit::commit_profiled` keeps the `performance-diagnostics`
  seam the crate-private native commit had.
- `Package` and `Presentation` stay on the `immutable_owners` rule.

## 4. Limits (§2.5)

- `raw::Reader::{open_with_options, from_reader_with_options,
  from_ole_file_with_options}`; the reader keeps the options' limits and
  leniency. `presentation_with_options` stays as the per-call credential
  form. `presentation_with_limits`, `presentation_with_options_and_limits`,
  `fonts_with_limits`, `picture_bullets_with_limits` are gone.
- `SourceBackedPackage::{from_read_at_with_options,
  from_shared_if_powerpoint_with_options, open, open_with_options}` (was
  `from_path`/`from_path_with_limits`), keeping its leniency; its
  presentation, picture-bullet and text-edit limit twins are gone.
- The other 70-odd `*_with_limits` twins of projections, builder drafts and
  record parsers had no consumer outside the crate and are crate-private; the
  public form applies the default limits. `odraw::ShapeExt` lost its limit
  twins (and `ppt_flags_with`) to a crate-private `ShapeLimitsExt`.
- `validate_source(source, PptValidationLimits)` takes its limits as a
  required argument.

## 5. Errors (§2.4)

- `slide_order::Refusal::kind()` and `text_edit::Refusal::kind()`; litchi-ops
  classifies both by kind (a protected kind keeps the signature suggestion).
  A transfer plan prepared for another receiver is
  `package::edit::Error::PatchConflict`.
- Code changes visible through litchi-ops: a duplicate shape-text target is
  `invalid_value` (was `invalid_argument`); a text-edit refusal for a shape
  without a textbox, anchor, text atom or notes page is `unsupported_type`
  (was `unsupported_capability`). The OfficeCLI corpus and capability goldens
  are unchanged.
- Scoped staging failures that the record grammar reported as free text
  (`InvalidFormat(String)`) are `Refused(InvalidValue)` with static
  operation/reason strings; publication failures of content the publisher
  cannot keep are `Refused(UnsupportedContent)`; absent targets
  `Refused(NotFound)`; unusable targets `Refused(UnsupportedTarget)`.

## 6. Fixes found on the way

Thirteen `Presentation` readers (PowerPoint 12 properties, routing slip,
embedded sounds, external media, OLE objects, modify password, privacy,
advisor, HTML, envelope data/settings, document structure, broadcasts)
scanned every `DocumentContainer` of the stream and refused any
incrementally saved presentation with "multiple Document containers". They
read the live document now, so read-after-edit works for every family. The
CRUD checklist therefore counts four PPT OLE instances in the corpus (was
three); it was re-blessed.

## 7. Performance

Routes of §5 for W3: PPT `set`/`add` (record 0323 `ppt_elements_probe`),
single set and batch/chunked text publication (record 0264
`ppt_paired_probe`). Method: release, no LTO, `taskset -c 29`, binary-level
ABBA (before, after, after, before), every run itself ABBA native/unified;
rounds: single 61, batch 15, elements 21. Before is trunk `5f6fb75164`
(`git archive`), after is this branch at the record commit. Binaries are not
committed; their SHA-256 are in
[`results/0357/binaries.sha256`](results/0357/binaries.sha256), the pooled
medians in [`results/0357/summary.json`](results/0357/summary.json).

| Probe | Scenario | Size | Native before → after µs | after/before | Unified before → after µs | after/before | Pair ratio before → after |
|---|---|---:|---:|---:|---:|---:|---:|
| batch | text batch (chunked, one publication) | 1000 | 3038.1 → 2986.7 | 0.983 | 3308.7 → 3246.7 | 0.981 | 1.090 → 1.088 |
| batch | text batch (chunked, one publication) | 4096 | 12337.1 → 12059.4 | 0.977 | 13847.8 → 13104.4 | 0.946 | 1.122 → 1.088 |
| batch | text batch (chunked, one publication) | 10000 | 32041.2 → 31707.2 | 0.990 | 36407.2 → 34517.7 | 0.948 | 1.138 → 1.088 |
| elements | ppt_cell_text_set_v1 | 1 | 104.7 → 104.3 | 0.997 | 108.6 → 108.6 | 1.001 | 1.038 → 1.041 |
| elements | ppt_cell_text_set_v1 | 50 | 391.0 → 391.7 | 1.002 | 401.4 → 400.5 | 0.998 | 1.017 → 1.022 |
| elements | ppt_comment_add_v1 | 1 | 106.5 → 106.3 | 0.998 | 109.2 → 109.2 | 1.000 | 1.027 → 1.028 |
| elements | ppt_comment_add_v1 | 50 | 446.1 → 446.3 | 1.000 | 452.1 → 450.9 | 0.997 | 1.019 → 1.012 |
| elements | ppt_comment_get_v1 | 1 | 1.3 → 1.3 | 0.992 | 1.8 → 1.8 | 1.000 | 1.415 → 1.438 |
| elements | ppt_comment_get_v1 | 50 | 1.3 → 1.3 | 0.992 | 1.8 → 1.8 | 0.994 | 1.433 → 1.439 |
| elements | ppt_comment_set_v1 | 1 | 102.7 → 102.9 | 1.002 | 107.5 → 108.3 | 1.007 | 1.047 → 1.050 |
| elements | ppt_comment_set_v1 | 50 | 359.4 → 358.8 | 0.999 | 370.2 → 370.7 | 1.002 | 1.029 → 1.029 |
| elements | ppt_hyperlink_add_v1 | 1 | 141.1 → 140.2 | 0.993 | 146.0 → 145.8 | 0.998 | 1.034 → 1.038 |
| elements | ppt_hyperlink_add_v1 | 50 | 518.6 → 513.8 | 0.991 | 527.7 → 524.8 | 0.995 | 1.014 → 1.022 |
| elements | ppt_run_format_v1 | 1 | 105.9 → 105.4 | 0.995 | 111.1 → 111.0 | 0.999 | 1.050 → 1.054 |
| elements | ppt_run_format_v1 | 50 | 401.2 → 398.4 | 0.993 | 412.0 → 411.3 | 0.998 | 1.033 → 1.030 |
| elements | ppt_run_get_v1 | 1 | 0.7 → 0.7 | 1.015 | 2.1 → 2.1 | 1.014 | 3.122 → 3.128 |
| elements | ppt_run_get_v1 | 50 | 0.7 → 0.7 | 1.019 | 2.1 → 2.1 | 1.001 | 3.133 → 3.074 |
| single | shape text set, warm | 1 | 103.6 → 103.5 | 0.999 | 107.9 → 107.9 | 1.000 | 1.041 → 1.040 |
| single | shape text set, warm | 100 | 129.8 → 129.2 | 0.995 | 136.0 → 135.7 | 0.998 | 1.040 → 1.042 |
| single | shape text set, warm | 1000 | 469.4 → 463.9 | 0.988 | 478.4 → 474.5 | 0.992 | 1.019 → 1.023 |

No scenario regressed: every after/before ratio is between 0.95 and 1.02,
and every pair ratio is within its budget (single-op write 1.10, batch 2.0)
except the two microsecond point reads (comment get 1.44, run get 3.1), which
are unchanged from record 0323 and carry its follow-up 0323-F12. The primary
edit's hot path did not change (the commit only gained a closure
indirection), the family plumbing is off the measured routes, and the
classification change is on the error path. Host load average 6.1 (other
agents' builds on other CPUs). The bench matrix gains the `0357` rows
(`tools/unified_bench_matrix.py` `p0357`).

## 8. Gates

Run on the branch after the final trunk merge (`03de00e6c4`):

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ppt -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`; with `--features ppt`; with `--features xlsb` | 0, 0, 0 |
| `cargo clippy -p litchi-ppt --no-default-features --all-targets -- -D warnings` | 0 |
| `cargo test -p litchi-ppt --all-features` (1521) | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib` (950), `--doc` (217), all 167 integration targets in six chunks | 0 (after re-blessing the CRUD checklist, §6) |
| `cargo test -p litchi --all-features` (209) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` | 0 |
| capability/registry goldens, all features and all-but-`encryption` (`LITCHI_BLESS=1`) | 0, 0 (no change) |
| `officecli_capability_delta.py --officecli …` (regenerated, no change) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ppt -p litchi-ops -p litchi` | 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (44 profiles, 2,949 bans, 277 families, 58 owners) | 0 |
| tool unit tests (the five required, `test_check_api_glossary`, `test_unified_bench_matrix`; 1,304) | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py` | 0 each |
| `check_crate_boundaries.py`, `non_iwork_gate.py verify` | 0, 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (301) | 0 |
| `tools/perf-baseline`: `cargo check --locked --all-targets`; `tools/native-resave`: `cargo check --all-features --locked` | 0, 0 |

Gate registrations (`ppt-minimal`, `ppt-full`): the 12 scoped families plus
the primary `package::edit` and `package::metadata` with `family_methods`;
crate-wide bans `rollback`, `is_changed`, `into_snapshot`, `Transaction`
(exact and suffix), `Editor` (exact and suffix), `_with_limits` (suffix);
the removed root aliases (`Font*`, `PictureBullet*`, `SlideSync*` staging
names, `Editor`, `EditorLimits`); the `get_*` options getters;
`presentation_with_options_and_limits`, `SourceBackedPackage::from_path`,
`writer::Chart`, `chart::{PackageEditor, Snapshot, edit_package,
into_package_editor, finish}` and `snapshot` on every family `Edit`.

## 9. Gaps and follow-ups

- Closed by [0357a §7](0357a-ppt-record-sized-patches.md#7-follow-up-restoring-the-routes-record-0357-made-crate-private):
  the slide-order durable patches, history, composition and three-way
  planning are `package::edit::Patch::{to_durable, join}`,
  `Package::{apply_durable, plan_three_way}` and inverse plus `apply`
  (§7.1). The source-backed same-length text splice is the registered
  `text_edit` family (`Package::{edit_text, apply_text}`), publishing a new
  `Package` over the composed overlay source without materializing it
  (§7.2). `document_structure`'s master reordering, custom-table-style
  placement and `fNonOutlineData` are `package::edit::Edit` verbs (§7.3).
- `client_data`-independent anchor sub-fields have no public route beyond
  the primary edit's verbs. Master-order, custom-table-style and
  `fNonOutlineData` changes have no durable patch vocabulary
  (`to_durable` refuses them).
- Closed by [0357a §6](0357a-ppt-record-sized-patches.md#6-follow-up-closing-the-0357-9-family-gaps):
  a reverted animation change publishing a changed package, the signed-deck
  refusal of `edit_animations()` and `edit_objects()`, `edit_objects()`
  refusing a deck without an external-object list, the master validator
  refusing real notes masters (13 corpus decks), and the missing
  title-master fixture.
- The scoped family patches carried whole-artifact bytes; record
  [0357a](0357a-ppt-record-sized-patches.md) made them record-sized
  byte-range diffs. `package::metadata::Patch` and `package::edit::Patch`
  still carry whole-artifact bytes.
- For W5: `litchi-ole-common`, `litchi-crypto` and `litchi-ograph` were not
  changed; none of their banned names reaches the PPT surface (the
  crate-wide bans pass).
