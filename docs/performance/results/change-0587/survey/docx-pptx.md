# Survey: litchi-docx and litchi-pptx facades (slug docx-pptx)

HEAD 2fc5fc657, read-only. Fresh evidence (this survey, scratch dir
`scratchpad/agents/docx-pptx/`): callgrind of three harness selectors at
`--warmup 0 --samples 1` (`incl-*.txt` = inclusive annotations, `native-*.json` =
one cold native sample each, `cg-*.log`), plus a native `perf record` self-symbol
list quoted inline. Callgrind runs use software SHA-256 (no SHA-NI under
valgrind), so SHA Ir shares are ~5x the native cycle share; native perf shares are
quoted next to them. All numbers are whole-child unless stated; the harness
corpora are synthetic.

## 1. Path map

**DOCX eager open** (`Package::open` → `from_opc_package`,
crates/litchi-docx/src/package/codec.rs:273-395): OpcPackage::open retains the
archive plus every decompressed part (0581); the facade adds only a content-type
check, `CustomProps::read_for` and `Slot::load` (core/app properties). Nothing of
document.xml is parsed at open; `mutable_doc` stays `None`.

**DOCX eager read** (`Package::document()` → `DocumentPart::from_part`,
crates/litchi-docx/src/parts/document_part.rs:576-584): every call runs the MCE
visibility pass (`visible_document_xml`, :159; copies the XML only when MCE
markup is present) and then `ParagraphIndex::from_xml` (:50), a full
`scan_word_element_ranges` pass, before any query. `Document::text()` →
`extract_word_text` (paragraph/codec/text.rs:39) is one pass into one `String`
(`try_reserve` per chunk); `write_text_to` (text.rs:422-470) streams to a sink
with no document-sized String but runs `preflight_semantic_xml` (text.rs:596) as
a second full pass. `paragraphs()` returns `Paragraph` views over one shared
`Arc<Vec<u8>>` (no per-paragraph copies). Source-backed `document()`
(source_backed.rs:610-660) does the same MCE pass + index build on the
materialized main part. Measured: `DocumentPart::from_part` 165.8 M Ir per 6
calls (index 114.5 M, MCE 51.2 M) versus `Document::text` 151.9 M per 6 calls —
the eager index alone costs ~75% of a full text extraction.

**DOCX ordinary edit + save** (`edit_document` → `Edit` → `commit` →
`publish_document_edit` → `to_stream`; package/package/document.rs:16-115,
document/transaction.rs:1834-1910, 4410-4460, 7241-7275): 
1. `document_snapshot` copies document.xml (`blob().to_vec()`) and
   `Snapshot::from_xml` scans paragraphs/tables/block controls (scan #1).
2. `replace_paragraph_text` rewrites one paragraph, `replace_range` copies the
   whole XML, `with_rewritten_xml` → `from_xml` rescans (scan #2), readback of
   one paragraph.
3. `commit` runs `compact_changed_document_xml` (writer/doc/codec.rs:45) over the
   whole main part, then rescans (scan #3).
4. `apply_document_patch` calls `document_snapshot` again (copy + scan #4) only to
   feed `Patch::apply`'s `same_source` byte compare (transaction.rs:1184, 4694),
   copies `candidate.xml_bytes().to_vec()`, then `edit_opc_candidate`
   (package/package/access.rs:203-245) clones the OpcPackage map (blobs are
   `Arc<Vec<u8>>`, so O(parts+rels) not O(bytes)) and re-parses core/app/custom
   properties.
5. `write_plain` → `PackageWriter` audits only authored/changed XML parts
   (pkgwriter.rs:123, `!package.is_exact_source_xml`), deflates them, and copies
   unchanged members when provenance exists (pkgwriter.rs:613-650).
Measured (`docx_semantic_one_edit_save`, shapes 24/200/10,000 paragraphs, summed):
`Snapshot::from_xml` 280.4 M Ir in 12 calls (4 per lifecycle) ≈ 64% of the
≈440 M timed Ir; compaction 62.6 M (≈14%); the actual one-paragraph rewrite +
readback ≈ 1 M; `validate_authored_xml` 93.8 M over 48 calls (content types,
rels, changed parts, shared with corpus creation); deflate 64.5 M. Native cold
single samples: 0.24 / 0.58 / 20.8 ms.

**DOCX managed/source-backed edit**: 0495-0500, 0517-0519 already cover it;
`commit` skips compaction for source-backed snapshots (transaction.rs:4415-4421).
Retained document-sized state on logical append: 0479-0485.

**PPTX eager open** (package/codec.rs:266-283): parses presentation.xml root
only (`PresentationPart::from_part`); no slide, master, layout or theme parse.
`Presentation::slide(i)` (presentation/package.rs:158-167) re-parses
presentation.xml (`slide_references`, parts/presentation.rs:118, NsReader + three
HashSets) on every call; `slide_count()` (:84) additionally resolves and
content-type-checks every slide part; `find_slide(Key::Name)` (:171) parses every
slide's XML. Measured: `Presentation::slide` 1.77 G Ir over 600 calls on a
200-slide deck = 2.95 M per call, proportional to N. `Presentation::text()` (:267)
builds per-slide Strings then concatenates; `write_text_to` (:45) streams per
slide (one bounded String per slide via `semantic_text_from_part`,
parts/slide.rs:808, which runs `scan_raw_semantic_text_xml` + MCE + parse = three
passes per slide). Source-backed open (`source_catalog`,
presentation/source.rs:3367-3385, `validate_slide_graph` :3317) reads
presentation.xml plus every slide's relationships (mandatory per 0577) but no
slide XML.

**PPTX opened transaction** (the ordinary opened-document CRUD route,
`opened_presentation_transaction` → `Transaction::commit` →
`apply_opened_presentation_commit`; opened/model.rs:308-395,
opened/transaction.rs:1195-1220, opened/patch.rs:753-870):
1. `capture_with_provenance`: `slides()` resolves every slide part, parses every
   slide's name (`c_sld_name` → MCE + parse), `notes::load_snapshot`
   (notes/package.rs:161-210) parses the root of every slide, then
   `package_fingerprint` (model.rs:397-440) SHA-256s every part blob including
   media, and `Arc::new(package.clone())`.
2. `Transaction::new` clones the OpcPackage map; edits touch only the selected
   slide (`Scene::read`, splice, `ensure_shape_id`).
3. `commit`: `compact_changed_slides` (changed slides only), then
   `package_fingerprint(&working)` to decide `unsign()`, `Patch::capture`
   (O(parts) with pointer-equal Arcs), then a full `capture` again (fingerprint
   #3).
4. `apply_opened_presentation_commit` discards `commit.snapshot`, calls
   `apply` (not `apply_exact_revision`), which clones the package, applies the
   deltas and runs `capture` again (fingerprint #4, notes scan, names).
Measured (`pptx_eager_batch_edit_save`, 200 slides × 8 text boxes + 8 × 2 MiB PNG,
39.6 G Ir whole child): `package_fingerprint` 13.84 G = 34.9% (16 calls ≈ 0.865 G
each: 13 via capture, 3 direct from commit); `capture_with_provenance` 13.19 G =
33.3% (13 calls: 3 open-transaction, 3 commit, 7 apply); `Transaction::commit`
≈2.0 G per call of which the two fingerprints are ≈86%; `apply_with_revision`
≈1.08 G per call, fingerprint ≈80%; `notes::load_snapshot` 1.41 G (108 M per
capture, `root_conformance` 1.33 G); `c_sld_name` 0.29 G. Native perf
(whole child): `x86_sha::compress` 10.8%, deflate ≈42% (harness corpus writes and
a provenance-less `from_bytes` save, see §4). Native cold lifecycle sample 342 ms.

**PPTX cross-package slide copy** (opened/cross_copy_plan.rs:880-920, 457-477,
523, 542-561, 614, 655, 1061, 1860-1880): planning computes
`physical_package_fingerprint` (a full `to_stream` serialization through a hashing
sink) of source and destination, `bounded_package_bytes` (a full serialization of
the candidate into a Vec), `Patch::capture`, then semantic and physical
fingerprints of the candidate; application recomputes semantic and physical
fingerprints of source and destination, physical of the candidate, and
`validate_candidate` → `capture` (semantic again). Measured
(`pptx_cross_copy_media_rich`, 166 G Ir whole child): SHA-256 79.5%;
`package_fingerprint` 77.2 G (46.5%, 44 calls via capture plus 21 direct);
`physical_package_fingerprint` 43.0 G (25.9%, 37 calls, each a full
`PackageWriter::write_to_stream`); `bounded_package_bytes` 21.7 G (13.1%, 8 calls,
re-deflating the candidate); `opened_presentation` 11.5 G (12 calls). Native perf
whole child: `x86_sha::compress` 28.0%, deflate ≈48%. Native cold lifecycle
sample 690 ms. Closure inventory itself (`collect_owned_closure`, `prove_*`,
`reject_mce`) is < 0.5%.

**Locks**: none on these paths except the `Mutex<Option<Error>>` failure slot in
the source-backed text sink (source_backed.rs:204-250). **Whole-package scans
without an index**: 94 `iter_parts()` loops followed by relationship-target
checks in litchi-pptx/litchi-docx production code (e.g.
`remove_unreferenced_dependencies`, opened/transaction.rs:1330-1375, O(parts ×
rels) per queued dependency); none appears in the profiled top symbols.

## 2. Remaining opportunities (ranked)

### 2.1 PPTX opened transaction: one complete-package revision, not four (step 1)
- Mechanism: per lifecycle the same package content is hashed four times
  (capture, commit unsign-check, commit recapture, apply recapture) and the
  notes/name scans run three times. Remove: (a) `commit` reuses the fingerprint it
  just computed for the recapture when `unsign()` changed nothing; (b)
  `apply_opened_presentation_commit` keeps `commit.snapshot` and uses
  `apply_exact_revision` + `validate_after` instead of re-deriving; (c) memoize
  per-part digests keyed by blob `Arc` identity + length inside `Snapshot` so a
  recapture hashes only parts whose Arc changed (revision = H(sorted per-part
  digests); bump the `litchi-pptx-opened-v1` domain string); (d) compute
  `SlideNameIndex` lazily (only `Key::Name` needs it) and cache the notes index
  per (presentation blob, slide blob) identity.
- Code: opened/model.rs:308-440; opened/transaction.rs:1195-1220;
  opened/patch.rs:753-870; package/model.rs:413-445.
- Records: none found for `package_fingerprint`/`capture_with_provenance`
  (grep of changes/*.md and top-level records). 0501 left 27.95% whole-child
  SHA-256 unattributed after removing `digest_touched`'s media input; 0449 assigned
  roughly half of SHA period to untimed harness output hashing. Not previously
  proposed.
- Size: measured 13.84 G Ir (34.9% whole child); modelled ≈3.46 G of ≈4.0 G
  Ir in open-transaction + commit + apply per lifecycle on the 16 MiB-media corpus.
  Native: SHA-NI 10.8% of whole-child cycles (upper bound for the timed region is
  unknown). On media-free decks the notes/name scans (≈130 M Ir per capture on 200
  slides) dominate instead.
- Scenarios: opened-document CRUD for PPTX (shape text, add picture, remove/move
  slide, same-package slide copy): `pptx_eager_batch_edit_save`,
  `pptx_eager_multi_slide_batch_edit_save`, `pptx_slide_*_boundary_save`,
  `pptx_cross_copy_*` (which build two Snapshots).
- Constraints: ADR 0003 (patch/revision binding) and ADR 0006 — the revision must
  still bind the complete package; (a)/(b) reuse a value computed on the identical
  `OpcPackage` instance; (c) changes the proof format and needs a frozen design
  record; (d) must keep ADR 0013 notes-topology checks before any mutation that
  touches notes. Risk: low for (a)/(b), medium for (c)/(d).
- Falsified if: native `perf stat` of the timed lifecycle on the media-rich and
  plain corpora shows the four passes below the 4% p50 noise floor after (a)+(b),
  or if a signature-policy test requires a post-unsign rehash that (a) skips.

### 2.2 PPTX cross-package slide copy: cache physical/semantic revisions (step 1-2)
- Mechanism: each copy serializes the whole source, destination and candidate
  through hashing or Vec sinks ≥5 times and hashes all blobs ≥6 times
  (cross_copy_plan.rs lines above). Cache `physical_package_fingerprint` per
  Snapshot (immutable `Arc<OpcPackage>`), compute the candidate's physical revision
  once with one hashing sink instead of `bounded_package_bytes` + a second
  serialization, and skip re-proving source/destination at apply when the
  Snapshot revisions already match the plan.
- Records: 0454 introduced the physical revision proof (not priced); 0501 scoped
  only the touched digest; 0423/0159 measured lifecycles. No record prices
  `physical_package_fingerprint` or `bounded_package_bytes`.
- Size: measured 141.9 G of 166 G Ir (85%) whole child under callgrind; native
  whole child SHA-NI 28.0% + deflate ≈48% (deflate includes corpus construction
  and provenance-less saves; the candidate serializations appear to re-deflate
  the 16 MiB of media — verify whether `build_candidate` loses provenance).
- Scenarios: `pptx_cross_copy_plain`, `pptx_cross_copy_media_rich` (+lifecycle),
  external-package copies (bin/pptx_external_cross_copy).
- Constraints: ADR 0003/0006 revision binding; "dedup only with proven
  equivalence" is unaffected. Risk medium; frozen design record required (proof
  reuse rules).
- Falsified if: with the plan+apply timers, caching leaves media-rich API p50
  within noise, i.e. the timed region does not contain these passes.

### 2.3 DOCX ordinary edit: stop rescanning and recompacting the whole document (step 1)
- Mechanism: (a) `apply_document_patch` builds a full Snapshot of the current
  main part just to compare bytes; compare `main.blob()` against
  `patch.before.xml_bytes()` (pointer or memcmp) and reuse `patch.after`; (b)
  after a same-structure paragraph rewrite, update the layout incrementally
  (shift ranges after the edit, grow enclosing table/control ranges) instead of
  `from_xml`; (c) compact only the replacement fragment before splicing, leaving
  untouched XML byte-identical, which also removes the post-compaction rescan.
- Code: package/package/document.rs:28-83; document/transaction.rs:937-943,
  1834-1910, 4410-4432, 7241-7275; writer/doc/codec.rs:45.
- Records: 0500 batched the *managed* scalar route (K=32 lifecycle 7x) and named
  repeated reconstruction as the hotspot; the ordinary batch route
  `replace_body_paragraph_texts` exists (transaction.rs:2504); 0518/0519 reuse
  proofs on the source-backed publication path. None cover the eager route's
  four scans or the whole-document compaction (grep `compact_changed_document_xml`,
  `apply_document_patch`: no record).
- Size: measured 280.4 M Ir of scans (≈64% of the ≈440 M timed Ir summed over
  the three shapes, large-dominated) and 62.6 M compaction (≈14%); the edit itself
  ≈1 M. Native: 20.8 ms cold for 10,000 paragraphs.
- Scenarios: opened-document CRUD DOCX (`docx_semantic_one_edit_save`,
  `docx_semantic_one_percent_edit_save`, `docx_semantic_noop_edit_save`),
  history undo/redo.
- Constraints: readback and `same_source` stay (ADR 0005 mandatory validation);
  (c) changes output bytes for documents whose untouched paragraphs contain
  compactable whitespace — arguably closer to ADR 0006 preservation-by-default,
  but it is a behaviour change needing a frozen record and preservation tests.
  Risk: low (a), medium (b, c).
- Falsified if: `docx_semantic_one_edit_save` large p50 improves < 4% after
  (a)+(b), or the layout scan turns out to be needed for the readback path.

### 2.4 DOCX eager `document()`: build the paragraph index lazily (step 1)
- Mechanism: `DocumentPart::from_part` always builds `ParagraphIndex` although
  `text()`, `write_text_to`, `tables()` and `blocks()` never use it; build it in a
  `OnceCell` on first `paragraph(i)`/`paragraph_count()`; the source-backed
  `document()` (source_backed.rs:625-645) has the same shape.
- Records: 0283 introduced the index for selected-paragraph queries; 0481/0482
  discuss the source-backed scanner index; none consider text-only reads.
- Size: measured 114.5 M Ir per 6 calls (≈19 M per call, 75% of one text
  extraction on the same document); MCE pass another 51.2 M (must stay).
- Scenarios: read/snapshot (`docx_file_eager_full_text`,
  `docx_semantic_full_text`, `docx_file_source_full_text`, `*_open_full_text_lifecycle`).
- Constraints: none (the index is already best-effort, `.ok()`-swallowed).
  Risk low; no ADR.
- Falsified if: the full-text harness timer starts after `document()` (then the
  gain is invisible to the selector, though real for callers).

### 2.5 PPTX eager facade: memoize the slide catalog (step 1/4)
- Mechanism: `Presentation::slide(i)` reparses presentation.xml per call
  (2.95 M Ir per call on 200 slides; index iteration is O(N²)); `slide_count()`
  resolves every slide part; `find_slide(Key::Name)` parses every slide. Cache
  `slide_references` in `PresentationPart` (borrowed, immutable) and let
  `slide_count` reuse the validated catalog.
- Code: presentation/package.rs:84-86, 158-211; parts/presentation.rs:118-200.
- Records: 0120 (root source path), 0375 (selected slide retained snapshot,
  source-backed); none propose caching on the eager view.
- Size: measured per-call cost above; whole-scenario impact unknown (the harness
  calls `slide(i)` once per sample).
- Scenarios: read/snapshot `pptx_file_eager_selected_slide`,
  `pptx_file_eager_slide_count`, by-name lookups. Risk low; no ADR.
- Falsified if: by-index iteration of a 200-slide deck via `slides()` is what
  callers use anyway and `slide(i)` p50 is already within noise of `slides()[i]`.

### 2.6 PPTX text extraction: three passes per slide (step 2, low)
- Mechanism: `semantic_text_from_part` runs `scan_raw_semantic_text_xml`, then
  MCE, then the text parse (parts/slide.rs:808-850); DOCX `write_text_to` runs
  `preflight_semantic_xml` before its parse. Fusing budget checks into the parse
  is the pattern 0514/0516 rejected for XLSX on measurement, so list as unmeasured.
- Size: unknown for the timed region (`text_from_part` 1.10 G whole child in the
  PPTX edit profile is harness verification). Falsified if a fused pass shows
  < 4% on `pptx_*_full_text`/`docx_*_full_text`.

## 3. Looks like an opportunity but is not
- Skipping `*/_rels/*.rels` reads at open: mandatory (0577, ADR 0005/0006).
- Routing `Package::open` to source-backed reads or fallible lazy `Part::blob`:
  0581 C2/C3 are frozen behind admission gates and each needs a proposed ADR.
- Dropping semantic readback or `same_source` checks in DOCX transactions, or the
  notes/relationship validation at PPTX capture: ADR 0005 mandatory validation
  (0574 opp 6 rejected the analogous early exit).
- Removing fdatasync from file-store publication: 0490 (79-82% of the tiny route
  is sync, "not justified" to remove).
- Byte-equal media dedup on slide copy: 0454 ("byte equality alone does not
  prove equivalent ownership").
- Rayon/global-pool parallel part reads for small local reads: 0498/0499 (worker
  waves regress small owned/file reads; only delayed providers benefit).
- A precomputed incoming-relationship index (Workstream E): 94 reverse scans exist
  but none shows in the profiled top symbols; maintaining it under OpcPackage
  mutation is ADR 0011 territory; no evidence yet that any scenario is bound by it.
- `OpcPackage` clones in `edit_opc_candidate`/`flush_presentation`/`capture`:
  blobs are `Arc<Vec<u8>>`, measured ≈1.1 M Ir per clone on 200 slides.
- Per-run `String` allocation in PPTX/DOCX text: DOCX already writes into one
  String or a sink; PPTX per-slide Strings are slide-bounded.

## 4. Measurement blockers and observations
- Eager PPTX harness corpora are built with `Package::from_bytes`
  (physical_source_provenance = false, package/codec.rs:252-258), so every eager
  save re-deflates all media (9.4 G Ir deflate in the edit profile, ≈42% native
  cycles); production `Package::open`/`from_vec` copy unchanged members
  (pkgwriter.rs:613-650). Eager save timings therefore overstate production cost;
  an `open`/`from_vec` variant of `pptx_eager_batch_edit_save` is needed.
- No selector times PPTX capture/commit/apply/publication separately (single
  `elapsed_ns`), and no eager-DOCX phase diagnostics exist (0496's overlay is the
  source-backed managed route), so the shares above are callgrind-modelled.
- Callgrind hides SHA-NI: SHA Ir shares are ≈5x native; native whole-child perf is
  quoted but no timed-region native attribution exists (needs a marker or
  `perf stat` around the API calls).
- Single native samples here are cold and uncontrolled (no A/A).
- 0501's replay verifier lacks its frozen control binary (GOAL_AUDIT 1090).
- Observation, not fixed: the ordinary DOCX `commit` compacts whitespace across
  untouched paragraphs (transaction.rs:4422-4429); check whether this is the
  intended preservation behaviour under ADR 0006.
- Observation: `apply_opened_presentation_commit` drops the commit snapshot and
  re-derives it with `apply` rather than `apply_exact_revision` (package/model.rs:413-418).
- Scratch retained: `incl-*.txt`, `native-*.json`, `cg-*.log`, `build.log`
  (772 KiB); callgrind `.out`, caller trees and perf data deleted after extraction.
  Normal harness binary rebuilt at tools/perf-baseline/target/release (44.8 s).
