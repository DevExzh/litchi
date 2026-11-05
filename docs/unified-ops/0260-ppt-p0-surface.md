# 0260 — PPT reaches the P0 read and write surface (plan step 13)

Status: implemented on branch `uo2/p13-ppt`, cut from `feat/unified-ops-api`
at `e467bcf4e8`. Plan: [0248](0248-coordinator-restart-plan.md) §6 Phase C,
step 13. Owned: `crates/litchi-ppt` (except its 0254 metadata module) and
`crates/litchi-ops/src/format/ppt*`. Shared dispatch files changed only in
small PPT arms (listed under "Shared files").

## What landed

### 1. Reads ported from `uo/m1` (L-PPT-R1..R4)

Commit `286487963d` was not on trunk (records 0210–0221 never touched these
readers). It was cherry-picked and adapted: `RecordLimits` became
`SourceLimits`, tests use the immutable owner's
`Package::open_with_options(.., OpenOptions::password)` and the shared
`PresentationData` fields.

- **R1, slide-list placeholder text.** A title/body placeholder whose
  `ClientTextbox` holds only an `OutlineTextRefAtom` now reads the slide-list
  body the reference selects (new public `OutlineTextBody`;
  `ShapeDefect::UnresolvedOutlineTextRef` for a dangling reference).
- **R2, one notes index** built at open from the parsed `DocumentContainer`.
- **R3, shape names and alternative text** (`wzName`, `wzDescription`):
  `ShapeEnum::{name, description}`.
- **R4, stable shape index**: `ShapeEnum::{source_index, id}`,
  `Slide::shape_by_id`.

Reviewed behaviour change: `text_shapes.ppt` and `41246-1.ppt` now include
placeholder text the reader used to drop. The facade Markdown golden
(`crates/litchi/tests/golden/markdown_corpus.tsv`, two rows, with a header
note) and the Python read test were re-recorded for exactly this change.

### 2. Shape text beyond same-length (record 0255's finding)

Record 0255 found PPT shape text writable only on `datetime.ppt` and only
at the same length. A corpus probe over every top-level shape of the 23
readable unencrypted PPT fixtures in `test-data` (214 targets, one
length-changing edit each) attributed the refusals: 104 `NoTextAtom`
(94 outline-backed placeholders, 10 bodies without a text atom), 85
`DependencyClosure`, 25 shapes without a textbox; 0 succeeded.

The native leaf now models:

| Record beside the text atom | Length change |
|---|---|
| `TextRulerAtom` (indents, tabs; no character position) | kept exact |
| `TextSpecialInfoAtom`, one run over the text (or text + 1) | count resized |
| `TextSpecialInfoAtom`, text run + one-character paragraph-mark run | first count resized, mark run exact |
| `TextSpecialInfoAtom` runs differing **only** in spell-check state | merged into one run marked unchecked (neither `error` nor `clean`), so PowerPoint re-checks; spelling is a proofing cache |
| `StyleTextPropAtom`, one paragraph run and one character run (optionally + mark run) | counts resized (was: single runs only) |
| outline-backed placeholder (`OutlineTextRefAtom`) | the selected `SlideListWithTextContainer` body is rewritten in the live `DocumentContainer`, in the same append-only publication |

Still refused for a **changed length** (same-length stays available), with
`Refusal::DependencyClosure`, because a whole-text replacement would
invalidate character positions that the binary format stores:
several paragraph or character formatting runs (which run would own the new
text is a semantic choice; an inverse durable replay could not restore the
run split), text interactions/hyperlinks (`InteractiveInfo`/`TxInteractiveInfoAtom`
ranges), metacharacter fields (`DateTimeMCAtom`, `SlideNumberMCAtom` at a
character position), and special-information runs that differ in language or
other text data. A body with no text atom (`NoTextAtom`) would need a new
atom inserted; a shape without a `ClientTextbox` has no text body.

`corpus_length_change_coverage_is_pinned` (litchi-ppt) pins the result:
**129 of 214 targets change length** (0 before), 47 closure refusals, 13
bodies without a text atom, 25 shapes without a textbox. Outline edits
reopen the published document, rebase staged slide structure over it at
commit, use the body for the signed no-op check, and replay through durable
patches. Independent check: LibreOffice 25 (`soffice --headless --convert-to
pdf`/`odp`) renders the edited title, outline body and drawing texts of
`SampleShow.ppt`, `41246-1.ppt` and `text_shapes.ppt`.

### 3. New native verbs

- `package::edit::Edit::set_notes_text(position, text)` (and on
  `slide_order::Transaction`): the notes page must hold exactly one
  `NotesBody` placeholder, whose body follows the closure rules above; the
  notes persist object is published with full readback. New
  `Refusal::NoNotesPage`; durable op `presentation.notes-text.set`.
- `Edit::insert_slide(position)` (named after PPTX `Edit::insert_slide`): a
  blank slide (`SL_Blank` = 0x10, following the master, color scheme and
  background of its neighbour; drawing with only the patriarch group and
  background shapes), as POI's `createSlide`. The drawing gets the next
  drawing identifier and a new shape-identifier cluster after every
  registered and already inserted drawing; commit appends one
  `OfficeArtIDCL` per new cluster and updates `OfficeArtFDGG`
  (`spidMax`, `cspSaved`, `cdgSaved`). Transferred slides that reuse a
  registered cluster are not registered again, so transfer output is
  unchanged. It shares the insert-transfer structural change, so exact
  patches, inverses, durable `slide-order.insert` and structural rebase
  cover it. An empty deck without a slide list gets one in MS-PPT document
  order. LibreOffice renders the added slides (page counts +1 on
  `SampleShow`, `41246-1`, `basic_test_ppt_file` and `empty.ppt`).

### 4. Unified PPT surface

| Element / verb | Before | Now |
|---|---|---|
| `/slide[N]/notes` | gap | get/query (walked after the shape tree, as PPTX), **set text** |
| `/slide[N]/shape[N]/paragraph[N]` | gap | get/query (text between `\r` marks; text shapes list paragraphs as children), **set text** (one native shape-text replacement per shape) |
| shape `name`, `alt`, raw `sourceIndex` | picture `name` only; `alt` gap | every shape kind; `@name` selectors on every kind; annotated/stats views report alt text (`picturesWithoutAlt`) |
| shape `text` | same-length, `datetime.ppt` only | length changes per §2 |
| `add` slide | unsupported | blank slides under `/`, every placement; `layout` may be `"blank"` |
| remove / move / swap slide | supported (0217) | unchanged |

Write contract, each with tests (`ppt_unified_text.rs`, `ppt_slide_add.rs`,
leaf `notes_text.rs`, `slide_insert.rs`, `shape_text_closure.rs`):

| Case | Behaviour |
|---|---|
| exact no-op (equal notes/shape text, empty props) | bytes and snapshot identity preserved, `patch.is_noop()` |
| signed source | changed write → `DocumentProtected`; native signed no-op keeps every byte |
| unlocked encrypted source | `DocumentProtected` (`EncryptionPolicy`) for set and add |
| released source (`retain_source(false)`) | `UnsupportedCapability` |
| read-back | every write reopens the published bytes; ops tests reopen through `File` too |
| history | exact `patch.apply` / `inverse().apply` restore source bytes |

Registry and capabilities: `notes` (text get/set) and `paragraph` (text
write-only key; text is the node member, as in PPTX) write tables; PPT `Add`
verb, slide `add`, notes/paragraph `set`; gaps for notes, paragraph and
picture alt removed; the shape-text gap names the remaining
position-dependent records. `ppt.json` capability golden re-blessed; the
OfficeCLI delta is unchanged (PPT is outside its matrix).

## Performance

Probe: `tools/perf-baseline/examples/ppt_paired_probe.rs` (release, no LTO,
`taskset -c 29`, ABBA rounds, one complete immutable write per sample,
open/setup/byte-equality checks outside the timer). Evidence:
[results/0260-ppt-paired.json](results/0260-ppt-paired.json) (raw samples,
binary SHA-256s; binaries not committed).

**Fix.** The single-set miss (0216: +6–39%) was one excess whole-slide parse:
the unified layer resolved `/slide[1]/shape[N]` through one slide handle,
then the native edit reparsed the same immutable bytes to resolve the same
target. Now `Presentation` keeps a one-slide memo of the last parsed slide's
shapes (keyed by persist id within the immutable projection; hits return
exactly what parsing would; failures are never stored; resident cost is one
slide's shapes, an ADR 0005 parsed-value cache bounded to one entry), and
`Package::edit` hands its own projection to the root transaction for target
resolution while the working artifact is still the source allocation.

Median unified/native pair ratios, before (`9da675a8fb`) → after
(`2bf3d6f7a9`), 21 rounds (7 for batches):

| Scenario | Before | After |
|---|---:|---:|
| single set, cold owner, 1 / 100 / 1000 shapes | 1.058 / 1.323 / 1.393 | 1.052 / 0.999 / 0.998 |
| single set, warm owner, 1 / 100 / 1000 shapes | 1.046 / 1.338 / 1.410 | 1.052 / 1.008 / 1.013 |
| 1k / 4096 / 10k text batch (one atomic edit) | 1.363 / 1.407 / 1.493 | 1.089 / 1.098 / 1.120 |
| outline title set (SampleShow) | 1.064 | 1.037 |
| outline body set (41246-1) | 1.014 | 1.010 |
| paragraph set (basic_test) | 1.048 | 1.042 |
| notes set (basic_test) | 1.058 | 1.063 |
| slide add (SampleShow / 41246-1) | 1.023 / 1.010 | 1.030 / 1.010 |

Every measured single-operation scenario is within the 10% budget; the 1k
batch is 1.09× (budget 2×). Absolute medians are not compared across the two
binaries: the host was shared (load 11–16) and repeated processes showed
bimodal absolute times (e.g. 276 vs 401 µs for the same native leg).

**10k scaling: still super-linear (not fixed).** Native 1k→10k growth is
35× (3.2 → 114 ms) and unified 36×. Attribution (native phase probe, 10k):
of 125 ms, 119 ms is staging; target resolution and read-back each parse the
whole 10 000-shape slide once per native edit (~45 ms each over the three
edits the 4096-operation hard bound requires). Each edit therefore costs
O(slide size), so a same-slide workload is O(operations × slide / 4096).
Making it linear needs either one publication beyond the 4096-operation
hard bound (a policy change, not a measurement fix) or a read-back that does
not reparse the whole slide (a validation change). Both are follow-ups for
the owner; neither was weakened here.

## ADR 0034 W4 residue (PPT)

Removed (not aliased) and banned in `ppt-minimal` and `ppt-full`:
`text_edit::SourceSnapshot::open_with_limits`, `rollback` on
`text_edit::{Transaction, SourceTransaction}` and `slide_order::Transaction`,
`slide_order::Snapshot::from_bytes_with_limits`. `ppt-full` gained the
`Builder::set_property` ban and the `package::metadata` family that 0254
registered only in `ppt-minimal`. The gate was checked to fail on a
re-added `slide_order::Transaction::rollback`. Glossary rows flipped:
Rollback, Configured open; new section "PPT P0 read and write surface".
Metadata §6 was completed by 0254.

Remaining PPT W4 residue: the scoped transaction families (`font`,
`client_anchor`, `client_data`, `external_media`, `broadcast`,
`document_comparison`, `diagram`, `master`, `master_layout`,
`picture_bullets`, `slide_sync`, `animation`, `property_set`) still expose
`Transaction`/`rollback`/`is_changed` and `*_with_limits` ingress, and
`raw::Reader::{open_with_limits, from_reader_with_limits}` is limits-only
raw ingress.

## Shared files

`crates/litchi-ops/src/write.rs` (PPT add arm), `transaction.rs` (PPT add
arms, `ppt` in the limits cfg), `capabilities.rs` (PPT `Add` verb, notes and
paragraph settable, slide addable, two PPT property-table arms, the
shape-text gap reason), `capabilities/tests.rs` and `tests/capabilities.rs`
(PPT expectations), `tests/value_allocations.rs` (warm native baseline),
`crates/litchi/tests/golden/markdown_corpus.tsv`, two pyo3 test files,
`tools/api_glossary.json`, `docs/API_GLOSSARY.md`,
`tools/perf-baseline/examples/ppt_paired_probe.rs` (new).

## Verification

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ppt -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-ppt --all-features` (1,473 tests) | 0 |
| `cargo test -p litchi-ops --all-features --no-fail-fast` (1,781) | 0 |
| `cargo test -p litchi --all-features --no-fail-fast` (260) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ppt -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` (67 packages, 270 declarations, 11 debt items) | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (21 profiles, 850 bans, 15 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (242 tests) | 0 |
| capability goldens (`LITCHI_BLESS=1`, ppt.json only) and `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI` (no change) | 0 |

No test or code references `3rdparty/`; every fixture is under `test-data`.

## Findings

- The builder writes `SlideLayoutType::Blank = 0x0D` in new decks; MS-PPT
  2.13.43 and POI define `SL_Blank = 0x10` (0x0D is unassigned). Opened-deck
  insertion writes 0x10. Changing the builder changes every built deck's
  bytes; left as a follow-up.
- `tools/perf-baseline` library code fails `clippy -D warnings`
  (`needless_borrow` in xlsx routes) independently of this change; the new
  example is clean.
- Record 0255's PPT row ("shape text only on `datetime.ppt`, add refused")
  is superseded by this record.

## Not claimed / follow-ups

- Layout-equivalent insertion (title/body placeholders) needs placeholder
  shapes, master placeholder references and slide-list bodies; only blank
  slides are created. Shape insertion/removal, nested group writes, rich
  text, run nodes and table cells remain gaps.
- Length changes over several formatting runs, interactions or fields
  (47 corpus targets) and bodies without a text atom (13) remain refused.
- Linear 10k batch scaling (see Performance).
- Scoped PPT transaction families (W4 residue above).
- Measurements are local probe results on a shared host, not frozen-harness
  qualification.
