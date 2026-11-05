# 0278 — ODP unified adapter: reads and writes over a scene index

- Status: implemented on `uo2/s23-odp`, cut from `feat/unified-ops-api` at
  `395d0ffa9f`. Plan [0248](0248-coordinator-restart-plan.md) §6 Phase F,
  step 23 (the ODP third of the ODT/ODS/ODP adapters), on top of the W3 ODF
  leaf wave of record [0268](0268-odf-w3-package-split.md).
- Date: 2026-10-13
- Owned: `crates/litchi-odp`, `crates/litchi-ops/src/format/odp*` and
  `crates/litchi-ops/src/format/odf_odp.rs`, plus the tests named below.
  One strictly additive item in `litchi-odf-common` (§4.3). Shared dispatch
  files changed in small, separated ODP arms only (§9).
- `performance_claim`: the paired leaf/unified measurement in §6 (release,
  one pinned CPU, ABBA × 12). No frozen-harness qualification.

Before this record ODP opened through the unified API but read only text,
counts and an always-empty metadata; `get`/`query` were refused with "M1 R7
adapter has not landed", and the capability report was a whole-format
refusal. ODP now reads and writes through the unified verbs.

## 1. Approach: a byte-range scene, edited by checked splices

The leaf's slide model (`litchi_odp::Slide`) flattens a slide into a title,
a body string and the non-text shapes, so it cannot address a paragraph, a
run or even a title frame by position, and its edit engine refuses title,
body and shape rewrites on every slide retained from a source package
(record 0002 §ODP, gap c). Rewriting retained slides from that model would
lose producer markup.

The adapter instead reads and writes through a new leaf index,
`litchi_odp::scene` (§4.1): one pass over `content.xml` records, for every
`draw:page`, its drawing objects in document order, their text stories
(paragraphs, including list items, and runs), tables, notes and the exact
byte range of each. Reads project values from the indexed producer markup;
writes become `scene::Change` values that `Edit::splice_scene` turns into
checked `XmlSplicePublication` edits on the primary `Package::edit()`
transaction of record 0268. Only the changed byte ranges of `content.xml`
are rewritten; every other byte and every other package member keeps its
source (the leaf's raw-member preservation applies to the rest of the
archive).

## 2. Read surface

Paths use the presentation vocabulary shared with PPTX and PPT.

| Element | Path | Source | Properties |
|---|---|---|---|
| presentation | `/` | package | `slideCount`; `title`, `author`, `subject`, `keywords`, `description` from `meta.xml` |
| slide | `/slide[N]` | `draw:page` | `name`, `slideNumber`, `title`, `layout`, `master` |
| shape (title, placeholder, text box, drawing shape) | `/slide[N]/shape[N]`, `/slide[N]/placeholder[TYPE]` | `draw:frame` with `draw:text-box`, `draw:rect`, `draw:custom-shape`, … | `name`, `alt`, `shapeType`, `placeholderType` (`presentation:class`), `isTitle`, `x`, `y`, `width`, `height`, `zorder`, `style`; raw `id` |
| picture / chart / ole / media | `/slide[N]/picture[N]`, … | frame with `draw:image` / `draw:object` (chart class) / `draw:object` / `draw:plugin` | `name`, `alt`, `source`, geometry, `zorder` |
| table, tr, tc | `/slide[N]/table[N]/tr[N]/tc[N]` | `table:table` in a frame | `rows`, `cols`, geometry, `style`; cells `colSpan`, `rowSpan` |
| group, connector | `/slide[N]/group[N]`, … | `draw:g`, `draw:line`/`draw:connector` | `name`, `alt`, `zorder` |
| paragraph | `…/paragraph[N]` under a shape, cell or notes | `text:p`/`text:h`, list items included | `style`, `level` (list depth) |
| run | `…/paragraph[N]/run[N]` | `text:span`, `text:a`, character data between them, text-bearing fields | `style`, `link` |
| notes | `/slide[N]/notes` | `presentation:notes` text frame | `text` |
| comment | `/slide[N]/comment[N]` | `office:annotation` on the page or its shapes (leaf inventory) | `text`, `author`, `initials`, `date` |
| slideMaster, slideLayout | `/slideMaster[N]`, `/slideLayout[N]` | `styles.xml` | `name`, `displayName`, `pageLayout` |

- Text follows the slide parser's whitespace rules (`text:s`, `text:tab`
  and `text:line-break` are explicit characters); a story's text is its
  paragraphs joined by `\n`.
- As in PPTX, a placeholder is a `shape` to selectors and a `placeholder` to
  paths; `title` and `textbox` narrow shapes through `isTitle`.
- Typed gaps (`unsupported_capability`, with follow-ups): video, audio
  (plugins are media), linebreak, hyperlink (links are runs), equation,
  notesMaster, noteSlide, animation, series, axis, col; property gaps
  `shape.fill`, `run.bold` (automatic styles are not resolved) and the
  `stats.fonts` view.
- Concepts ODF lacks are `unsupported_path`: smartart, zoom,
  modernComment, reply, model3d, theme.
- Views: text, annotated, outline and stats (OfficeCLI shapes), the others
  are the shared view gaps.
- `File::metadata()` now projects `meta.xml` (it returned an empty value).
- Flat presentations (`.fodp`) open read-only through the leaf's
  `FlatPresentation` (previously refused); the leaf still refuses a DOCTYPE
  and a drawing-bodied document with a typed `corrupt_file`.
- Master pages are listed by identity (`master::list`, §4.2): the shared
  semantic master reader refuses LibreOffice ≥ 7.6 master pages (their
  `loext:theme` child), which made two corpus files' masters unreadable.

The query-corpus differential (`tests/query_corpus_differential.rs`) gained
37 ODP selectors over the 11 `.odp` and 8 `.fodp` files in `test-data`
(76 digests). They were blessed from the read-only build (commit
`341933bad4`) before any write handler landed, and reproduced byte for byte
by a second run then and after the write commits (release, 31 s).

## 3. Write surface

| Verb | Elements | Leaf change |
|---|---|---|
| set `text` | shape/placeholder story, notes, table cell | `SetStoryText`: the story becomes one paragraph per line in the first paragraph's style and list wrappers (identities dropped); a self-closing container is expanded |
| set `text` | paragraph | `SetParagraphText`: character data when plain, else the whole paragraph re-authored around its own start tag |
| set `text` | run (span, link, plain text) | `SetRunText`; field runs refuse |
| set `name`, `x`, `y`, `width`, `height` | shapes, frames, tables (geometry), groups/connectors (`name`) | `SetAttribute`: one rewritten start tag per element |
| set `name` | slide | `SetPageName`; refused when another page has the name or links/custom shows reference the old one |
| set `/metadata` | `title`, `author`, `subject`, `description`, `keywords` | `Edit::set_metadata`: one aligned element splice of `meta.xml` per changed field |
| add | slide under `/` (optional `name`, any placement); text box under a slide (`text`, `name`, `x`, `y`, required `width`, `height`; appended); paragraph under a shape, notes or cell (`text`, any placement) | `InsertPage` (copies the neighbour's page style, master and layout), `InsertTextBox`, `InsertParagraph` (a list item when the neighbour is in a list) |
| remove | slides (the last one stays), drawing objects, paragraphs (with a list item or list they alone fill), runs | `RemovePage` / `RemoveObject` (refused when its id is referenced) / `RemoveParagraph` / `RemoveRun` |
| move, swap | slides; top-level objects of one slide (z-order) | `MovePage`, `SwapPages`, `MoveObject`, `SwapObjects`: the element's own bytes are relocated (§4.3) |
| create | presentation with one blank slide; initial metadata limited to the five fields | `Builder` |
| save, write_to | atomic replacement of the origin or an explicit path, `Overwrite::Refuse` honoured | `litchi_core::atomic` |

- **Exact no-ops.** A change whose value equals the current one plans
  nothing; a call or transaction that changes nothing returns the source
  snapshot, whose bytes are the input's (`exact_noops_publish_the_source_bytes`).
- **Atomic and immutable.** Every verb and transaction leaves its input
  untouched; a failed transaction names its item and publishes nothing.
- **Coalescing.** In `File::edit` and `batch`, consecutive direct
  positional-path sets join one pending run that publishes as one
  `splice_scene` call on one leaf edit and one commit. A later set of the
  same text target replaces the pending value; attribute sets merge into
  their start tags; a set that overlaps a pending one at another
  granularity (a story and its paragraph, a paragraph and its run) or
  renames a page twice closes the run first. Queries, identity paths,
  `/metadata` and structural verbs take the ordinary route. Overlap checks
  are hash lookups, so a run is linear in its length.
- **Signed and encrypted packages** are refused by the leaf edit, as in
  record 0268; flat documents report `retainsSource: false` and refuse the
  edit verbs with a typed `unsupported_capability`.

## 4. Leaf changes

### 4.1 `litchi-odp` (commits `2e04dc55dc`, `90662f3537`, and the write commit)

- `scene::{Scene, Page, Notes, Object, ObjectKind, FrameContent, Story,
  Paragraph, Run, RunKind, Table, Row, Cell}` (read index) and
  `scene::{Change, ObjectRef, Scope, StoryRef, Attribute, NewTextBox,
  NewPage}` (edits). Planning resolves every change of one call against the
  scene of the content staged before the call and refuses overlaps.
- `Edit::{scene, splice_scene, metadata, set_metadata}` on the primary
  transaction (new `content::Operation::{SpliceScene, SetMetadata}`
  replay arms).
- A commit whose only domain is these checked splices does not re-audit
  `content.xml` and `meta.xml` as whole compact authored parts: their
  authored fragments were audited at staging and every other byte is
  source-proven. Without this, any change to a producer part with a
  newline after its XML declaration (every LibreOffice file) was refused at
  commit. Other domains keep the full audit.
- `annotation::inventory(content_xml)`, `master::{parse, list, Summary}`:
  source-free reads for the source-backed and flat owners; `master::list`
  reads identities only and so tolerates `loext:theme`.

### 4.2 Metadata patch alignment

`patch_meta_xml` returns a whole new part (and drops a byte-order mark).
`set_metadata` publishes each changed field separately, aligning the
difference to whole elements (`element_splice`) so the authored fragment
is one balanced element; the shared single-diff derivation
(`xml_splice_publication`) started fragments mid-tag on insertions.

### 4.3 `litchi-odf-common`: one additive item

`AuthoredXmlFragment::relocated(&XmlSourcePart, Range)` — balanced markup
copied from another range of the same source part keeps the producer's
spelling and is not held to the compact authored contract (the assembled
part is still verified). Slide and object moves and swaps need it: their
bytes are source bytes, and LibreOffice writes `<x a="b" />`, which the
authored audit refuses. New test file
`crates/litchi-odf-common/tests/xml_splice_relocation.rs`. Nothing else in
`litchi-odf-common` changed.

## 5. Capabilities, registry and delta table

- `Capabilities::of(Odp)` supports get/find/query/view, set/add/remove/move/
  swap/edit/batch, save and create; element operations and property flags
  come from the ODP write registry (`format/odp/registry.rs`); ten operation
  gaps list the refused contexts (field runs, character formatting,
  non-paragraph stories, positioned text-box insertion, picture/table/chart
  creation, paragraph reordering, referenced pages and objects, the
  set-only `/metadata` path, extended metadata).
- New goldens `tests/golden/capabilities/odp.json` and
  `tests/golden/registry/odp.json`, blessed with `--all-features` and with
  every feature except `encryption`; both builds wrote identical files and
  no other golden changed.
- `docs/unified-ops/officecli-capability-delta.md` regenerated with
  `python3 tools/officecli_capability_delta.py --officecli
  /home/zhuhe/code/litchi/3rdparty/OfficeCLI`: the tool already compares
  ODP against the PPTX vocabulary (OfficeCLI has no ODP schema), so the
  only change is the new ODP section; `--check` and
  `python3 -m unittest tools.test_officecli_capability_delta` pass.

## 6. Measurements

`crates/litchi-ops/examples/odp_unified_bench.rs`, release, `taskset -c 7`
(AMD EPYC 9R45, shared host), 12 rounds alternating leaf/unified (A B B A),
medians of per-call means. The deck is authored through the unified API:
50 slides of one 20-paragraph text box (1,000 paragraphs, 10,000 for the
last row: 100 × 100). The leaf leg is the hand-written equivalent:
`Package::from_bytes` + `Scene::parse` for reads, one `Package::edit()` +
one `splice_scene` + `commit()` for writes; every write row first asserts
both legs publish identical bytes. Raw rows:
`docs/unified-ops/results/0278/odp_unified_bench.tsv`.

| Scenario | leaf µs | unified µs | overhead | bar |
|---|---:|---:|---:|---|
| open + first get | 640.1 | 647.2 | +1.1% | open ≤ 5% ✅ |
| query `paragraph` (1,000 matches, warm scene) | 0.5 | 69.6 | n/a | streaming iterator ✅; see below |
| single set (one paragraph) | 1771.2 | 1775.6 | +0.3% | ≤ 10% ✅ |
| batch 1 | 1769.6 | 1778.1 | +0.5% | |
| batch 10 | 1805.6 | 1822.3 | +0.9% | |
| batch 100 | 1856.5 | 1953.2 | +5.2% | |
| batch 1,000 | 2339.1 | 3212.3 | +37.3% (1.37×) | ≤ 2× ✅ |
| batch 10,000 | 29248.4 | 37895.6 | +29.6% (1.30×) | linear ✅ |

- **Two misses found and fixed while measuring.** The first build ran the
  coalescer's overlap check as a scan of the pending run (quadratic: batch
  10,000 was 8.5× the leaf) and admitted every write by building and
  dropping a whole leaf `Edit` (+29% on a single set). Both are fixed: the
  overlap check is a hash index, and admission reads the package's
  security policy (the leaf edit is built only to report a refusal).
- **The query row has no leaf counterpart.** The leaf leg is a bare loop
  over borrowed scene paragraphs; the unified leg builds a typed node per
  match (its path steps and address), about 70 ns each. No budget applies
  to query cost per match; `first()` and `limit()` stop the walk.
- **Most of a write is the leaf edit.** `Package::edit()` builds the slide
  staging draft (a full slide parse) even when only content splices
  follow; that is about 1.7 ms of the 1.77 ms single set on this deck and
  is the same on both legs (follow-up G-ODP-6).

## 7. CLI and Python

- `litchi-cli`: `tests/cli.rs` reads an ODP fixture in the `READS` table,
  creates an ODP in the creatable-format loop, and adds a text box, sets
  its text, swaps and moves slides and removes one, each saved in place.
- `pyo3-litchi`: `OdpVerbTests.test_round_trip` (set, add, swap, move,
  remove, `/metadata`, reopen, `save_as`); ODP left the read-only set and
  joined `CREATABLE`; the flat-ODF message snapshot of the drawing-bodied
  `.fodp` now records the leaf reader's refusal.

## 8. Gaps and follow-ups

- **G-ODP-1 character formatting.** Bold, size, colour and fonts live in
  automatic styles; reads report style names and writes author plain text
  in the first paragraph's style. Follow-up: resolve and author automatic
  text styles.
- **G-ODP-2 creation of non-text objects.** Pictures, tables, charts and
  media need package resources the scene splice does not author; only text
  boxes are added. Positioned (non-append) text-box insertion is refused.
- **G-ODP-3 references.** Removing a page or an object whose name or id
  other markup references, and renaming a referenced page, are refused
  rather than rewriting the references.
- **G-ODP-4 paragraph/run/row/cell ordering** is not offered (slides and
  top-level objects move and swap).
- **G-ODP-5 shared master reader.** `litchi-odf-common`'s
  `style::master` reader refuses `loext:theme`; the adapter lists masters
  by identity instead. The fix belongs to the ODF common owner.
- **G-ODP-6 leaf edit cost.** `Package::edit()` parses every slide into the
  staging draft even for content-only edits; making the draft lazy would
  cut the leaf (and unified) write cost on large decks.
- **G-ODP-7 flat writes.** `.fodp` documents are read-only; their scene
  splices would apply to the single XML document, but the flat leaf has no
  splice publication yet.
- **G-ODP-8 find/replace and template merge.** After the trunk merge
  (records 0273, 0274), `set` with `find=` (`Verb::Replace`) and
  `Verb::Merge` are refused for ODP with `unsupported_capability`; the
  scene's run text changes can carry a find adapter.
- **0268 follow-ups G2–G4** were not needed by this adapter and remain
  open; G5 (security vocabulary) is unchanged.
- **Duplication for later consolidation.** `format/odf_odp.rs` (`/metadata`
  assignment of the five ODF fields, ODF length spelling) is ODF-generic;
  the ODT and ODS adapters landing concurrently carry their own copies.

## 9. Shared files

Small, separated ODP arms only: `read.rs`, `write.rs`, `transaction.rs`,
`save.rs`, `create.rs`, `capabilities.rs`, `node.rs`, `node/view.rs`,
`query/{matches,plan}.rs`, `view/{json,presentation}.rs`, `open.rs`,
`open/{bytes,path}.rs`, `format/mod.rs`; tests `tests/capabilities.rs`,
`tests/registry.rs`, `tests/query_corpus_differential.rs`,
`tests/support/query_corpus.rs`, `crates/litchi-cli/tests/cli.rs`,
`crates/pyo3-litchi/tests/{test_write_api,test_errors}.py`. No iWork crate,
`Cargo.toml`, or `PROGRESS.md` changed.

## 10. Commits

| Commit | Content |
|---|---|
| `2e04dc55dc` | leaf scene index, checked splices, metadata edit |
| `90662f3537` | leaf source-free annotation/master reads, page masters |
| `341933bad4` | unified read adapter; query-corpus digests blessed here |
| `930c1b1bf4` | `AuthoredXmlFragment::relocated` (odf-common, additive) |
| `626576d352` | leaf splice fixes found by the write adapter |
| (this record) | unified write adapter, goldens, delta, CLI, Python, bench |

## 11. Gates (run on the final tree, all exit 0)

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-odf-common -p litchi-odp -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo test --no-fail-fast -p litchi-odf-common -p litchi-odp -p litchi-cli --all-features` | 0 (1,023 passed) |
| `cargo test --no-fail-fast -p litchi-ops --all-features` | 0 (1,981 passed; includes the query-corpus differential) |
| `cargo test --no-fail-fast -p litchi --all-features` | 0 (260 passed) |
| `cargo test --release -p litchi-ops --all-features --test query_corpus_differential` | 0 (before and after the write commits) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-odf-common -p litchi-odp -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir <wt>/target/glossary` | 0 (26 profiles, 1,201 bans, 38 families) |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (253 tests) |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry`, with `--all-features` and with every feature but `encryption` | 0; only `odp.json` goldens written, identical in both builds |
| `python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI`, then `--check` and `python3 -m unittest tools.test_officecli_capability_delta` | 0, 0, 0 |

## 12. Trunk merge (`fcc4fe5266`)

Trunk's records 0273–0277 (find/replace, merge, RTF, ODT and ODS adapters)
were merged into this branch. Every ODS, ODT and ODP dispatch arm, cfg list
entry, coalescing run and `wrap_*` helper is kept. The query-corpus golden is
the union of both sides' digests (no existing key re-recorded); its selector
arrays carry per-entry `cfg` attributes. Python: ODS, ODT and ODP are all
creatable, `READ_ONLY` is empty, and the flat ODF message rows cover `.fodt`
plus the flat ODP leaf refusal. The capability and registry goldens were
re-blessed with and without `encryption` (the encryption-free build also
rewrote trunk's stale `capabilities/docx.json` and `capabilities/xlsx.json`,
which lacked the `replace` verb), and the delta table was regenerated.
