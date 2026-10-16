# 0303 — PPTX/DOCX structural gaps on files from other producers

- Status: implemented on `uo2/pptx-gaps`, cut from `feat/unified-ops-api`
  at `771a38c752`, trunk merged at `e7fe30fb24`. Follows record
  [0293](0293-producer-structure.md) §4 gaps 1, 2 and 5, and record
  [0290](0290-checklist-failures-and-metadata-language.md) §3 (F3).
- Date: 2026-10-16
- Owned: `litchi-pptx` (slide removal, the source-backed editor and its
  durable wire, the slide guard), one additive `litchi-opc` API, the PPTX
  capability-gap text in `litchi-ops`, and the `crud_matrix`,
  `corpus_fidelity` and `producer_structure` tests. `litchi-docx` is
  unchanged (§5).
- `performance_claim`: none. §7 shows no touched path got slower beyond
  noise.

## 1. Result

| # | gap | outcome |
|---|---|---|
| 1 | PPTX slide that owns notes, pictures, diagrams, VML, comments or OLE objects | **removed with the Parts it owns exclusively**, reversibly, on the source-backed and the eager editor (§2) |
| 2 | PPTX `p14:sectionLst` | **kept consistent** on source slide remove, move, swap and add (§3) |
| 3 | F3: PPTX OLE remove under markup compatibility | **fixed**: whole `p:graphicFrame` removal on the source bytes; the guard stays for every other edit (§4). F3 leaves `KNOWN_FAILURES` |
| 4 | DOCX drawings in run-level `mc:AlternateContent` | **not admitted**; §5 records exactly which proofs are missing, with an experiment |

Structure census (`structure_census`, PPTX, 80 decks, 69 with one slide):

| attempt | 0293 after | 0303 |
|---|---|---|
| remove first slide | 1 | 9 |
| remove last slide | 3 | 9 |
| slide move + back (11 decks) | 8 | 10 |

The two decks still refused: POI `bug62513.pptx` (its ZIP has trailing
bytes, which the OPC topology publisher refuses; not a PPTX question) and
`smartart-autoTxRot.pptx` (§2.4).

## 2. Slide removal with owned Parts

### 2.1 Orphan proof

`presentation::owned_parts` builds the typed relationship graph of the
package (the source, a prepared topology candidate, or an eager OPC
package) and computes, for a set of removed slides:

1. every Part reachable from them through internal edges, except Parts of
   the presentation structure (slides, layouts, masters, notes and handout
   masters, themes, the presentation itself), which a slide never owns;
2. the greatest subset of those in which every incoming edge comes from a
   removed slide or from the subset; the package root counts as an outside
   owner.

Each Part in the result is an orphan once the slides leave; a Part
reached but not kept (a layout, the notes master, a picture another slide
shows) stays. A notes slide's edge back to its slide is inside the set.
Joint removal matters: a picture shared by two slides removed in one
edit leaves with them, and a durable replay proves ownership for all the
patch's removed slides together.

A slide with no internal edge besides its layout and slide jumps owns
nothing, so the package-wide graph is not built (`may_own` fast path).

The incoming-owner proof of the slide itself now skips owners in the
owned set; any other owner (a slide jump from another slide, a custom
show binding) is still a typed `SharedOwner` refusal.

### 2.2 Source-backed editor

The topology recipe (`topology_wire::Data`) gains `dependents`: one
`Dependent` per owned Part with its exact payload and relationship images
before and after. It composes like slide changes (recipes are relative to
the physical source; a Part restored and removed again drops out), is
diffed in `patch_from_packages`, rebased on apply, and compared in
`same_content` and the physical-target check.

- **Wire.** A patch with dependents is encoded as version 5 (`LPSX0005`,
  own checksum domain; `LPSX0004` is the shape-removal envelope). A patch
  without dependents is still version 3, byte for byte. Decoding checks
  that dependents are sorted by folded name, unique, distinct from changed
  slides, and that each changes presence.
- **Apply.** A removal patch must remove exactly the Parts its removed
  slides own on the current topology; a restore must restore exactly the
  Parts its restored slides own on the prepared candidate. Any other
  dependent set is a `PatchConflict`.
- **Restore.** A restored Part gets its typed edges, then its exact
  relationship image. Restored XML (slides and dependents) is held to the
  source XML policy through the new
  `litchi_opc::SourceTopologyPlan::try_add_restored_part_shared`; before,
  a durable inverse after reopening failed on any producer slide with
  indentation (`NotCompact`), so only same-process inverses worked.
- **Refused up front.** An owned Part whose internal edges are not spelled
  owner-relative, and any edge of the slide or its owned Parts to a missing
  Part: a typed edge cannot restore either after reopening.

The source plan (`SourceBackedPresentationEditor::plan_slide_removal`)
and the eager planner (`Snapshot::plan_slide_removal`, whose durable
`SlideRemovalPatch` now admits extra removed non-slide resources) remove
the same closure.

### 2.3 Lenient registration reads

Removal reads only the layout registration. `validate_registered_removal_layout`
reads the presentation root and masters leniently: root extensions and
extension namespaces, master extension namespaces (`a14:hiddenFill`),
master background-picture relationships and themes with relationships no
longer refuse. The presentation's slide bindings are still proven by the
removal planner. Slide insertion's layout lookup reads the same way.

### 2.4 Still refused

- a slide another slide jumps to (`SharedOwner`), as before;
- a slide whose relationships name a missing Part (`smartart-autoTxRot`
  names a `diagramDrawing` its producer never wrote): removal would be
  safe, but its durable inverse could not restore the edge after
  reopening;
- the final slide, signed, macro-enabled and protected packages, as
  before.

## 3. Sections

`opened::sections` gains `remove`, `strip` and `validate_membership`, and
the source-backed editor uses the eager editor's placement rules:

- **remove**: the slide's `p14:sldId` entry leaves with its catalog entry.
  A section that loses its last slide stays; PowerPoint keeps empty
  sections. The source plan publishes the root as two checked lexical
  deletions.
- **move**: the moved slide joins the section of its new neighbour; a swap
  keeps membership and re-sorts sections. A sectioned order is derived
  from the current root, so a sequence of moves composes, and a slide
  moved out and back across a boundary may keep its new section: the root
  then changes without an order change.
- **add**: the inserted slide joins its neighbour's section.

Proofs: the topology root-delta proof and the `LPSX0002` order patches
strip section entries before their byte proofs and check the membership
(same section count, no slide in two sections, entries in presentation
order, retained slides keep a section). The order-invariant root digest
leaves membership out; the exact root preimages guard it. The standalone
`SourceBackedSlideOrderEdit`, which rebuilds each candidate from the
source, still refuses sections.

## 4. F3: OLE removal under markup compatibility

`transition::admit_source` classifies a slide as `SlideMce::Plain` (the
old guard passes) or `SlideMce::Enclosed`: every markup-compatibility
element lies strictly inside a top-level `p:spTree` child, the rest of the
transition guard holds, and no namespace is bound by character
references. Anything else is refused as before. Snapshots carry the class.

On an Enclosed slide:

- **Removal** (`opened::stage_enclosed_shape_removal`): the shape is
  chosen and its connector/animation dependencies are checked on the
  processed scene; the processed shape-tree children must be exactly the
  scene roots, and the source must have as many children, so the k-th
  root is the k-th source child; that child's source bytes are deleted;
  the candidate's shapes must be the source's without the removed ones,
  in order and by identity.
- **Other edits** (text, geometry, text properties, text boxes, slide
  name/visibility, transitions, SVG) are refused with the guard's
  `UnsafeEdit`.
- **Replay** (`durable::validate_pair` under the shape-removal envelope)
  proves an Enclosed pair as whole-child deletion with no retained
  reference to a removed shape; restored slides and publication snapshots
  use the same admission.

A byte comparison of processed forms was tried first and dropped: the MCE
processor re-serializes the whole part when it rewrites anything (the
newline after the XML declaration goes), so processed bytes of the source
and of the MCE-free candidate are never comparable.

## 5. DOCX run-level markup compatibility: not admitted

The fixture is `libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx`:
one `w:r` holds an `mc:AlternateContent` whose Choice is a `wps` text box
and whose Fallback is VML, each with a `w:txbxContent` paragraph
(`<Text>` in both here). The empty-edit probe still fails with
`unsafe DOCX edit 'edit' rejected: source-backed document transactions do
not support markup-compatibility branch selection`.

Two barriers, established by a throwaway experiment that forced each one
open (not committed):

1. **Admission.** `litchi_ooxml_common::mce::admission::admit` refuses any
   element in the MCE namespace (`Refusal::Element`), and the projection
   is rewritten, so `admit_document_source` returns `UnsafeEdit`.
2. **Positions.** `VisiblePayload::project` marks `RawPositions::Unproven`
   whenever the projection reports an `AlternateContent`. With admission
   forced, every body `set`/`remove` fails with "the DOCX block position
   is not proven in the raw part" and `add` fails.

With both forced, body paragraph add and remove publish the expected
shifted documents, and the reader's paragraph text for the drawing's
paragraph includes the Choice text box text (`<Text>…`). So the work left
is three proofs, not a new engine:

- an admission variant proving every `AlternateContent` is a direct child
  of a `w:r` (no block-level alternate content), so the body block and run
  positions of the source and the projection agree;
- a position verdict that accepts that case;
- the raw transaction scanners (the fast and the `NsReader` oracle scanner)
  treating those runs as opaque, and refusing any edit that addresses that
  run's paragraph text or run content, since the projection shows one
  branch's text-box text as the paragraph's.

Follow-up: "0303 DOCX run-level markup-compatibility admission".

## 6. Checklist and fidelity

- **CRUD checklist** (`results/0289/crud-checklist.{md,tsv}`, re-blessed):
  pptx `ole` `remove` is `ok` on `ole.pptx` `/slide[1]/ole[1]`; pptx FAIL
  1 → 0. With trunk's record 0298 (F8), `KNOWN_FAILURES` is now empty.
  Items 1 and 2 do not show here: the checklist's remove/move cells run on
  documents Litchi made.
- **Fidelity** (`results/0289/fidelity.{md,tsv}`, re-blessed after the
  trunk merge, whose record 0294 adds one PPTX fixture): PPTX structure
  byte-exact 9 → 11 (both sectioned LibreOffice decks), refused 3 → 1
  (`bug62513`). `tdf103477`'s text probe is still refused, now by
  the Enclosed edit guard at `set_text_properties` instead of at the
  snapshot. No other row changed; every no-op save stays byte-identical.
- **Capabilities**: the `remove slide dependency-topology` gap describes
  the owned-Part removal and what it still refuses; the slide-ordering
  gaps no longer list sections. Goldens re-blessed in both feature sets;
  `officecli-capability-delta.md` regenerated (reason text only).

Tests:
- `litchi-pptx`: `tests/producer_slide_removal.rs` (pictures, notes,
  comments + VML + OLE, shared and jointly owned pictures, slide-jump
  refusal, missing-Part refusal, eager editor; each with the local inverse,
  the transported forward patch and the durable inverse after reopening),
  `tests/producer_slide_sections.rs` (remove, empty section, move, move
  out and back, swap, composition, insertion, eager removal),
  `tests/producer_ole_removal.rs` (exact frame bytes, inverses, refused
  text and slide-property edits, a plain shape beside the frame), wire
  round trip and validation tests in `topology_wire`; the eager plan test
  and `source_backed_slide_removal` now expect the owned Part removed;
  `slide_sections::source_insertion_places_the_new_slide_in_a_section`
  replaces the pinned section refusal.
- `litchi-ops`: `tests/pptx_producer_gaps.rs`; `producer_structure`'s notes
  refusal became a removal.

## 7. Measurements

`examples/pptx_gaps_bench.rs`, release, pinned to core 7, 21 rounds after
3 warm-ups, legs alternated ABBA inside each binary, binaries run
before/after/after/before twice. "before" is the trunk tip `8cf481f5fd`
plus the example; "after" is this branch. Evidence:
[`results/0303/pptx-gaps-bench-abba.jsonl`](results/0303/pptx-gaps-bench-abba.jsonl).
Medians in µs (run 1 / run 2).

Unified vs native (after):

| scenario | native | unified | unified/native |
|---|---|---|---|
| remove slide 1 of `45545_Comment.pptx` (notes, comments, VML, OLE) | 2404.5 / 2412.6 | 2413.9 / 2420.2 | 1.004 / 1.003 |
| remove the OLE frame of `ole.pptx` | 177.3 / 177.5 | 178.6 / 179.1 | 1.007 / 1.009 |
| move a slide of `slide-sections.pptx` | 225.7 / 226.9 | 226.2 / 228.2 | 1.002 / 1.006 |

The base refuses all three.

Before/after on changed plain paths (`shapes.pptx`), after/before:

| scenario | native | unified |
|---|---|---|
| remove dependency-free slide 6 | 0.973 / 0.966 | 0.973 / 0.964 |
| move slide 1 after slide 2 | 1.004 / 1.003 | 1.006 / 1.009 |
| set one shape's text | 0.995 / 0.983 | 0.995 / 0.989 |

- **Slide removal** is ~3 % faster: removal no longer runs the copy-surface
  validator over the presentation root and master XML (§2.3). This is a
  side effect, not an optimization claim.
- **Move**: the first after-build was 1.009–1.016, from the section probe
  that now runs on every order change. Switching it to `memchr::memmem`
  (memchr becomes a regular `litchi-pptx` dependency; it was already in
  the tree) brought it to 1.003–1.009. callgrind of the move scenario:
  4,772,263,719 → 4,778,954,462 instructions (+0.14 %).
- **Shape text set**: unchanged (the snapshot admission is the old guard
  on plain slides).

Binaries are not committed. SHA-256: before
`85396e0ee26798d0a98424b35363fd339083778b3af795b1b2da69d291d02b12`,
after `169f942e2fe6a35c499303b63ffb2550fa35eb960a1fd1284e213f8a2586eadc`.
Build: `cargo build --release -p litchi-ops --features pptx --example
pptx_gaps_bench`. Host load 4–6 on 32 cores.

## 8. Gates

All exit 0 on the final tree (after the trunk merge), with
`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8` and one cargo command at a
time.

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-opc -p litchi-pptx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features pptx\|docx] --all-targets -- -D warnings` | 0 each |
| `cargo test -p litchi-opc -p litchi-pptx --all-features --no-fail-fast` | 0 |
| `cargo test -p litchi-ops --all-features` (`--lib`, `--doc`, then the 133 integration tests in chunks of 25; rerun after the second trunk merge) | 0 (one chunk first failed in `pptx_slide_remove`, which pinned the notes refusal; updated, then 0) |
| `cargo test -p litchi --all-features --no-fail-fast` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-opc -p litchi-pptx -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `python3 tools/test_no_3rdparty_references.py`; `python3 tools/unified_bench_matrix.py --check` | 0 each |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (279 tests after the second trunk merge) |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 (lock gains `memchr` under `litchi-pptx`) |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry`, `--all-features` and every feature but `encryption` | reviewed (§6) |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI>` | 0, reason text only |
| `LITCHI_BLESS=1 … --test crud_matrix` / `--test corpus_fidelity` | reviewed (§6) |
| `cargo test -p litchi-cli --all-features --test officecli_corpus`; `python3 tools/officecli_corpus.py --check` | 0, 0 (unchanged) |

## 9. New or changed items

- `litchi_opc::SourceTopologyPlan::try_add_restored_part_shared` (public,
  additive).
- `litchi-pptx`: crate-private `presentation::owned_parts`,
  `opened::enclosed_removal`, `transition::{SlideMce, admit_source}`,
  `opened::sections::{remove, strip, validate_membership}`; topology wire
  version 5. Public signatures are unchanged; the public removal planners
  admit more (documented on `remove_slide` and `plan_slide_removal`).
- The glossary gate passes; no glossary row changes.

## 10. Not claimed

- The census removes one slide per position (first, last); other slides
  may meet other constructs.
- The enclosed-MCE admission is proven for MCE inside top-level shape-tree
  children only. MCE wrapping a whole top-level shape
  (`mc:AlternateContent` as a `p:spTree` child) is still refused.
- Restored Parts' ZIP entries are recompressed; the Part bytes, their
  relationship XML and `[Content_Types].xml` are exact, not the ZIP
  container bytes.
