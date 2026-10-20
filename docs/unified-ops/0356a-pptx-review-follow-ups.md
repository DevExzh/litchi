# 0356a — PPTX review follow-ups: source-backed families, no raw OPC mutators, selectors

- Status: implemented on `uo2/w2-pptx` after the coordinator review of
  record [0356](0356-pptx-api-contract.md) (tip `02f405c299`, merged into
  trunk as `5e55079b38`); trunk merged at `184dd4d908` (records 0357, 0353a,
  0356) and at `7893492f9e` (record 0355a); for coordinator review.
- Date: 2026-10-20.
- Input: the review's decisions P1–P4 on 0356's gaps 0356-P1 to 0356-P4, and
  its two acceptances (the per-part conflict check and the unchanged
  encryption mode after saving).
- Owned: as 0356 (`crates/litchi-pptx/**`, the PPTX hunks of `litchi-ops`,
  the facade and `tools/perf-baseline`, the `pptx-*` profiles, line W2), plus
  `tools/api_debt_inventory.py` and `results/0351/inventory.tsv` (P4).
- `performance_claim`: measured (§4). Open, `set`, the 1k batches,
  `to_bytes` and `to_bytes` twice are within −2.6%…+0.3% of trunk with 0356
  (after/before, A-B-B-A); every 1k batch route is within 1.97× of the leaf.

## 1. Result

| Decision | Done |
|---|---|
| P1: source-backed stream editors get the family shape on the source-backed owner, no capability loss | `source::Package::edit()` was already the source-backed slide, multi-slide and order family; its `Commit` is now the core triad alias and the module carries `Error`, `LimitKind`, `Diagnostics`. The SVG editors become `source::svg` (`Package::edit_svg_images`: `replace`, `attach`, `detach`, `image`, `attachment`; `apply_svg_images`). The 20 `SourceBacked{Slide, SlideBatch, SlideOrder, Svg, SvgAttachment}{Edit, Commit, Patch, Snapshot}` types and the editor's `edit_*`/`publish_*_to_stream` routes are crate-private and banned |
| P2: no public function mutates a raw `OpcPackage` | 0 (was 63 definitions; checked over the rustdoc JSON, `.coord/w2-pptx/scripts/rawmut.py`). 11 new families; the raw twins of 8 existing ones are crate-private |
| P3: families take the API's selectors | `slide::{Selector, MasterId, LayoutId, Owner}`, `Position` for an item on a slide, typed XML IDs (author, comment, section, slide ID); package-level parts are documented (§2.3) |
| P4: the inventory counts a core `Commit` alias as having the accessors | `tools/api_debt_inventory.py`; `results/0351/inventory.tsv` regenerated (§5) |

Commits: `83892956f8` (P3, P4 tool), `8d766e76fd` (P2), `2a1aaa45e8`
(P1), this record's commit (evidence, matrix, glossary line, inventory);
trunk merges `efbdebe2de` and `2764481df0`.

## 2. What landed

### 2.1 P1: source-backed families

The bounded-memory route is kept. `source::Package` holds the retained
source and only the changed slide XML; a family commit returns a new
`source::Package` whose untouched parts are still lazy references to the
source, and writing it (`write_to`, `to_bytes`, `save`) raw-copies every
untouched member, so a commit followed by a save streams exactly as the
former `publish_*_to_stream` routes did.

- **Slide, multi-slide and order edits** were already one family,
  `source::Package::edit() -> source::Edit` (every verb of the stream
  editors: slide name, visibility, properties, transition; shape text,
  properties, removal; text properties and their per-edit change flags;
  text boxes; paragraph and run insertion and removal; slide insertion,
  removal, move and swap) with `source::{Patch, Commit}` and
  `Package::apply`. `source::Commit` becomes
  `litchi_core::edit::Commit<source::Package, source::Patch>` and
  `source::Diagnostics` the core one. The one read only the stream editor's
  slide snapshot offered, the slide transition, is
  `source::Slide::transition()`.
- **SVG pictures** (`source::svg`). An SVG change is not a slide overlay:
  it retargets media parts and the slide's relationships. Its commit is a
  *deferred* `source::Package` (the shape of DOCX record 0355a): the commit
  writes nothing; its output is the stream route's topology publication,
  run when the package is written, and the first read or edit of the
  committed package materializes that output once as a source-backed
  package over the published bytes (`primary/deferred.rs`). Verbs:
  `replace(slide, image, svg::Replacement)`,
  `attach(slide, image, &svg::AttachmentReplacement)`, `detach(slide,
  image)`; reads `image(slide, image) -> svg::Image` and `attachment(slide,
  image) -> svg::Attachment`. A changed edit of a signed source is refused
  at the verb. An SVG patch is exact-source: it replays onto a package over
  the same retained source, and its inverse applied to the commit's own
  package returns the package the commit was made on (the 0355a `restore`
  pair).
- The SVG value types move into the family: `svg::{Image, Attachment,
  AttachedSvg, Replacement, AttachmentReplacement}` (were
  `SourceBackedSvgSnapshot`, `SourceBackedSvgAttachmentSnapshot`,
  `SourceSvgAttachment`, `SourceSvgReplacement`,
  `SourceSvgAttachmentReplacement`).
- The stream editors' code stays in the crate: the slide and SVG snapshot,
  edit and patch types are the engine `source::Package` and `source::svg`
  run on; the batch and order editors and the editor's own publication
  entry points are kept `#[cfg(test)]` for the engine tests, which move from
  `tests/` into `src/presentation/source/engine_tests/` (the routes they pin
  are crate-private now). `SourceBackedPresentationEditor` itself stays
  public for its cross-slide-copy and slide-removal plan routes, which are
  not edit families.
- Two family tests cover the public shape: `source::svg` replace (the
  committed package's bytes equal the stream route's, the source is
  unchanged, reads see the change, the inverse restores it, a stale replay
  conflicts, a no-op keeps the bytes, a signed source is refused) and
  attach/detach.

### 2.2 P2: raw `OpcPackage` mutators

New families, generated from the raw functions they replace
(`.coord/w2-pptx/scripts/raw_family.py`) and documented per verb:

| Family | `Package::edit_x` | Verbs |
|---|---|---|
| `comments` (legacy) | `edit_comments` | `comments`, `author`, `comment`, `set`, `add_author`, `update_author`, `remove_author`, `reorder_authors`, `add`, `update`, `remove`, `reorder` |
| `modern_comments` | `edit_modern_comments` | `comments`, `graph`, `authors`, `set`, `set_authors`, `comment`, `add`, `update`, `replace`, `remove`, `reorder`, `reply`, `add_reply`, `update_reply`, `replace_reply`, `remove_reply`, `extensions`, `update_extensions`, `reply_extensions`, `update_reply_extensions`, `author`, `add_author`, `update_author`, `replace_author`, `remove_author`, `reorder_authors`, `changes`, `update_changes` |
| `shape::theme` | `edit_themes` | `theme`, `summary`, `theme_override`, `presentation_theme`, `add`, `attach`, `set_colors`, `set_fonts`, `set_override`, `remove_override` |
| `presentation::embedded::vba` | `edit_vba` | `project`, `set`, `remove` |
| `presentation::embedded::ink` | `edit_ink` | `add` |
| `presentation::embedded::ole` | `edit_ole_objects` | `add` |
| `media_parts` | `edit_media` | `media`, `set` |
| `chart` | `edit_charts` | `add` |
| `laser` | `edit_laser_traces` | `set` |
| `presentation_properties::metadata::revision` | `edit_revision_info` | `revision_info`, `set` |
| `presentation_properties::metadata::slide_sync` | `edit_slide_sync` | `slide_sync`, `set` |

Crate-private twins of existing families: `font::{put, remove}`,
`tag::{put, remove}`, `tag::shape::{put, remove}`,
`table::style::{put, remove}`, `read_only_recommended::{put, remove}`,
`presentation_properties::{put_math_to_package,
remove_math_from_package}`, `master_layout::{add_slide_master,
add_slide_layout, remove_slide_layout, store_placeholder_shape}`,
`designer_tags::{store, remove}`. Deleted as aliases:
`replace_presentation_comment`, `replace_presentation_comment_author`
(equal to `update_*`), the `store_slide_media` re-export. Read helpers over
`&OpcPackage` stay public. The raw escape is `Package::edit_opc`.

### 2.3 P3: selectors

`slide::MasterId` and `slide::LayoutId` are the typed `p:sldMasterId/@id`
and `p:sldLayoutId/@id` (unique in a presentation); `slide::Owner` is a
slide (`Selector`), a layout or a master; `SlideMaster::id()` and
`SlideLayout::id()` read them. Converted signatures:

- `master_layout::Edit::add_slide_layout(MasterId, ..)`,
  `remove_slide_layout(LayoutId)`, `set_placeholder_shape(impl Into<Owner>,
  ..)`; `placeholder::Edit::{set, remove}` and
  `Package::placeholder_type_extension` take an owner.
- `color_map::Edit::{update, color_map}` take an owner.
- Slide selectors for `events::Edit::{update, set, remove, show_events}`,
  `ink_actions::Edit::{ink_actions, edit_profile, replace_profile,
  replace_profile_bytes}`, `content_parts::Edit::update`,
  `ole::slide::Edit::update`, `controls::slide::Edit::update` (the control
  by `Position`), `threading::Edit::{threading, set, remove}`,
  `notes::Edit::{set_text, set_slide_xml}` (were notes-graph indexes),
  caption tracks (`tracks::Edit::{update, media}`: slide and the picture's
  shape ID, which `MediaKey` held as a part name) and every new P2 family
  verb that took a slide part name.
- Package-level parts keep no selector: the presentation part's families
  (guides, custom shows, structure, Changes and Revision Information, VBA,
  slide synchronization) take none; themes are named by part
  (`/ppt/theme/themeN.xml`) because a theme is shared by masters, notes and
  handout masters, and an added theme has no owner until it is attached.
  The theme module documents this.

Test fixtures that build orphan parts outside the presentation lists keep
their raw route through crate-private test seams.

### 2.4 P4: inventory

`api_debt_inventory.py` reads the gate's foreign-path map: a type whose
caller-visible path aliases `litchi_core::edit::Commit` has the triad
accessors (they are checked by the core's unit tests). The PPTX rows left
are the ten `&mut self` methods of the `builder::Presentation` draft
(drafts keep `&mut self`, 0351 §2.1 rule 7; the inventory's owner list
names every `Presentation`). The regenerated
`results/0351/inventory.tsv` covers the 25 crates and `litchi-ops` of 0351
§1 on this branch's tree (trunk `7893492f9e` merged, record 0355a
included): 216,674 bytes, SHA-256
`19f315dde216931d280662d8f655a8a7c1240a407db3f1384ecec9fde3cbf6fd`
(0351 cites the trunk `5ff605692e` file). Regenerate with
`.coord/w2-pptx/scripts/inventory.sh` (the 0351 docstring commands, one
crate at a time).

## 3. Behaviour (accepted in review, recorded here)

- **Per-part conflict check.** A family `Patch` (the part-level engine of
  0356 §2.2) conflicts when a part it changes does not hold its exact
  before bytes, content type and relationships; a part it does not change
  may differ. The detached snapshot revision check still guards each
  `update` against its own staged graph. Before 0356 the detached
  transactions compared a revision of their whole snapshot. Accepted by the
  coordinator review of 0356.
- **Saving does not change the encryption mode.** `save_encrypted`,
  `to_encrypted` and the re-encrypting outputs take `&self` and leave the
  package's retained encryption provenance as it was opened; before 0356 a
  save marked the package with the mode it wrote. Accepted by the review.
- Selector strings select slides by producer-visible name: a slide part
  name passed where a selector is expected is a name, not a part
  (`"/ppt/slides/slide1.xml"` names a slide called that). The P3 migration
  found one test relying on the old meaning.
- `perf-baseline`'s source-backed one-edit, batch and multi-slide benches
  now measure `source::Package::edit` and `write_to` (the stream editor
  routes are crate-private); their evidence is not comparable with earlier
  runs of the same names.

## 4. Measurements

Release, `taskset -c 7`, shared host. Evidence in
[`results/0356a/`](results/0356a/) (116 KB), binaries and commands in
`results/0356a/binaries.sha256`; "before" is trunk `184dd4d908` (0356
merged), built from the same `pptx_contract_bench` source.

| route (fixture) | before µs | after µs | after/before |
|---|---:|---:|---:|
| open, leaf (`45545_Comment.pptx`) | 1,638.1 | 1,617.1 | 0.987 |
| open, unified (same) | 730.4 | 732.4 | 1.003 |
| `to_bytes`, leaf (same) | 12,109.9 | 12,122.9 | 1.001 |
| open + `to_bytes` twice, leaf (same) | 25,925.7 | 25,940.3 | 1.001 |
| open + `to_bytes` twice, unified (same) | 746.2 | 746.3 | 1.000 |
| open, leaf (20-box deck) | 277.5 | 278.4 | 1.003 |
| `to_bytes`, leaf (20-box deck) | 711.8 | 706.4 | 0.992 |
| `set` shape text, unified | 477.2 | 466.0 | 0.977 |
| `set` + `to_bytes`, unified | 552.1 | 542.7 | 0.983 |
| 1k run-format batch, unified | 14,304.8 | 14,081.7 | 0.984 |
| 1k shape-text batch, unified | 10,734.3 | 10,460.5 | 0.974 |

Record 0348's routes, three runs (leaf against unified, in the after build):
every 1k route within 1.97× — cell batches 1.06–1.33, paragraph text
1.12–1.25, run format 1.13–1.29, shape text 1.02–1.09, notes cold
1.91–1.94 and warm 1.19–1.20; every 1k and 10k route within the 2.0
budget, the worst being 10k cold notes at 1.954–1.986 (0356: 1.945–1.972,
0348: up to 1.966; the leaf and unified code of these routes is unchanged
since 0356). Single operations 1.00–1.09×. Every scenario checks equal
leaf and unified bytes (all equal). The matrix rows supersede 0356's.

## 5. Gates

On the tree merged with trunk `7893492f9e` (merge `2764481df0`); logs in
`.coord/w2-pptx/gate-logs/` (`gates-summary.txt`). The workspace commands
exclude only `litchi-ods` (its purged evidence fixture, as in 0356).

| gate | exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-pptx -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features pptx] --all-targets -- -D warnings` | 0, 0 |
| `cargo clippy -p litchi-ops --features <all but encryption> --all-targets -- -D warnings` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pptx -p litchi-ops` | 0 |
| `cargo test -p litchi-pptx --all-features` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, 167 `--test` targets in chunks of 20 (executables deleted per chunk) | 0, 0, 0 |
| `cargo test -p litchi --all-features`; `cargo test -p litchi-cli --all-features` | 0; 0 |
| `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify` | 0; 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (44 profiles, 433 families, 66 owners) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (301 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_unified_bench_matrix tools.test_check_api_glossary` | 0 |
| `tools/officecli_capability_delta.py --check`; `tools/officecli_corpus.py --check` | 0; 0 |
| `tools/unified_bench_matrix.py --check` | 0 |
| `tools/check_feature_matrix.py` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry` (all features; all but `encryption`), then a clean tree | 0, 0, 0 (no golden changed) |

The OfficeCLI corpus and batch-equivalence tests run in the `litchi-ops`
chunks. Target directory after cleanup: 17 GB (limit 20 GB).

## 6. Follow-ups

- 0356a-P1: `SourceBackedPresentationEditor` keeps its public
  cross-slide-copy and slide-removal plan routes
  (`SourceBackedCrossSlideCopyPlan`, `SourceBackedSlideRemovalPlan`,
  `publish_cross_slide_copy_to_stream`, used by the perf-baseline
  retention benches). They are plans, not edit families, and outside this
  review's P1; moving them onto `source::Package` would retire the editor
  owner.
- 0356a-P2: `api_debt_inventory.py` could classify `builder::*` draft
  types as drafts rather than owners, so that "no rows left" is literal for
  PPTX.
