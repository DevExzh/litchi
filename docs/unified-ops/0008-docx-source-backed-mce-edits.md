# 0008 — DOCX source-backed edits and markup compatibility: diagnosis and options

- Milestone: M2 (write kernel). Blocker raised by M0a S5a
  (`uo/m0a`: `docs/unified-ops/m0a-log.md:29`, test
  `crates/litchi-ops/src/format/docx/tests.rs:400-429`).
- Status: diagnosis and option analysis. **No production code changed.** The
  recommendation in §5 awaits the coordinator; stages 1–3 change leaf
  behaviour and should land as their own records.
- Base: `uo/0008-docx-mce` at `cae4526597`. Paths below are relative to the
  repository root; line numbers are at that commit.
- `performance_claim: none`. Timings are indicative single-host medians from a
  throwaway probe, not registered claims.

## Answers in brief

1. **Which MCE constructs trigger the refusal.** None of the semantic ones.
   The main part of every one of the 19 `test-data/ooxml/docx/*.docx`
   fixtures has no `mc:AlternateContent`, `mc:ProcessContent`,
   `mc:MustUnderstand`, `mc:PreserveElements` or `mc:PreserveAttributes`.
   13 fixtures (Word 2007 and Mac Word 2008) only *declare* the MCE namespace
   (`xmlns:ve="…/markup-compatibility/2006"` on `w:document`) and never use
   it. The other 6 carry exactly one `mc:Ignorable` on `w:document`, and the
   MCE projection drops nothing but that attribute. The refusal fires because
   the MCE namespace URI occurs in the bytes. It does not fire because the
   projection changes anything the editor reads.
2. **Root cause.** The guard compares allocations (`Arc::ptr_eq`). The MCE
   codec returns the input allocation only when the namespace URI is absent,
   and otherwise always re-serializes. The invariant the guard defends is
   real: the transaction's coordinates and published bytes are the raw part,
   the reader's view is the projection, and the two must agree on everything
   an edit reads or writes. The check is just much stronger than that
   invariant, and in one case (an entity-escaped namespace, W7 below) it is
   also weaker.
3. **Eager owner.** It edits the raw part with no MCE check at all. On every
   fixture its edit changes only the target paragraph (or inserts one
   paragraph), keeps every `mc` construct and every other ZIP member byte for
   byte, and produces exactly the bytes the source-backed publisher produces
   once its guard is bypassed. It also accepts the edits the guard exists to
   stop (wrong-target positions under body-level `AlternateContent`, a
   `MustUnderstand` the profile cannot satisfy), and it publishes into three
   protected fixtures that the source-backed route refuses.
4. **Options.** See §4. Briefly: (a) is correct but large and needed by no
   fixture yet. (b) is small, low-risk and unlocks all 19. (c) works but loses
   either physical preservation (`to_owned_package`) or resident-memory
   economy, and adds no safety. (d) stages (b) and extends it to both owners
   and to the managed route.
5. **Recommendation.** Replace the pointer test with an admission proof, in
   stages. Stage 1 admits declared-only documents (13/19). Stage 2 admits
   Ignorable-only documents (all 19). Stage 3 applies the same admission to
   the eager owner and to the managed route. Stage 4 defers run-level
   `AlternateContent` to a PPTX-style mapped edit with a readback proof. Do not
   fall back silently to the eager owner.

## 1. Which constructs the fixtures carry

### Method

A throwaway probe (§7) classified each fixture's main part, resolved through
`_rels/.rels`, with `lxml`. It then drove the leaf APIs:
`source_backed::Package::{edit_document, document, document_variables_snapshot}`
unmanaged and under an `ExecutionContext`;
`litchi_ooxml_common::mce::process_markup_compatibility` with the
`visible_document_xml` profile (OOXML baseline plus `w14` and `w16du`,
`crates/litchi-docx/src/parts/document_part.rs:498-517`); and the eager
`litchi_docx::Package`.

### The 19 fixtures

| Fixture | Producer | MCE use in `word/document.xml` | What the projection drops | raw → visible bytes | Doc-vars today |
|---|---|---|---|---|---|
| `ComplexNumberedLists.docx` | Word 2010 | `mc:Ignorable="w14 wp14"` on `w:document` | the `mc:Ignorable` attribute | 4,948 → 5,368 | refused |
| `FancyFoot.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 2,891 → 3,005 | ok |
| `Headers.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 2,213 → 2,416 | ok |
| `Hyperlink.docx` | Word 2013 | `mc:Ignorable="w14 w15 wp14"` | the `mc:Ignorable` attribute | 1,862 → 1,910 | refused |
| `comment.docx` | LibreOffice 6.0 | `mc:Ignorable="w14 wp14"` | the `mc:Ignorable` attribute | 1,264 → 1,336 | ok |
| `documentProperties.docx` | Mac Word 2008 | declares `xmlns:ve`, never uses it | nothing | 1,319 → 1,366 | ok |
| `documentProtection_comments_no_password.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 1,036 → 1,070 | ok |
| `documentProtection_forms_no_password.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 1,041 → 1,075 | ok |
| `documentProtection_no_protection.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 1,025 → 1,059 | ok |
| `documentProtection_readonly_no_password.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 1,036 → 1,070 | ok |
| `drawing.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 288,070 → 318,958 | ok |
| `endnotes.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 4,906 → 5,162 | ok |
| `footnotes.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 1,360 → 1,457 | ok |
| `headerFooter.docx` | Mac Word 2008 | declares `xmlns:ve`, never uses it | nothing | 1,352 → 1,494 | ok |
| `headerPic.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 1,170 → 1,250 | ok |
| `table-alignment.docx` | Word 2016+ | `mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"` | the `mc:Ignorable` attribute (156 `w14` attributes are understood and kept) | 40,151 → 42,053 | refused |
| `table-indent.docx` | Word 2013 | `mc:Ignorable="w14 wp14 w15"` | the `mc:Ignorable` attribute | 71,991 → 79,157 | ok |
| `table_footnotes.docx` | Word 2007 | declares `xmlns:ve`, never uses it | nothing | 2,478 → 2,682 | ok |
| `testComment.docx` | Word 2016+ | `mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14"` | the `mc:Ignorable` attribute (2 `w14` attributes kept) | 3,108 → 3,181 | refused |

**Distribution.** 13 declare the namespace and use nothing. 6 have one
root `mc:Ignorable` whose only effect is its own removal. There are 0
`AlternateContent`, 0 `ProcessContent`, 0 `MustUnderstand`, 0 `Preserve*`, and
0 attributes or elements in an ignorable namespace that the profile does not
understand. The codec's `Report` agrees on all 19: `alternate_content_count`,
`ignored_elements` and `unwrapped_elements` are 0, and `ignored_attributes` is
0 for the 13 and 1 for the 6 (the directive itself). All 19 refuse
`edit_document()` today with
`UnsafeEdit { reason: "source-backed document transactions do not support markup-compatibility branch selection" }`,
and all 19 refuse even `document()` under an execution context (§2.4).

For all 19, the visible buffer equals the raw bytes after exactly three
lexical rewrites, checked byte for byte:

- whitespace between the XML declaration and the root is not copied;
- on the 6, the root start tag is rebuilt without `mc:Ignorable`, with single
  spaces between attributes;
- every empty tag `<q …/>` becomes `<q …></q>`.

Nothing else differs, so no paragraph, run, text or attribute that the editor
reads changes.

Two corrections to the S5a note: 6 of the 19 carry `mc:Ignorable`, not
"every fixture", and `comment.docx` is LibreOffice output, not Word.

### The wider corpus (63 DOCX-family files under `test-data/`)

| Main-part MCE shape | Files |
|---|---:|
| no MCE namespace | 1 (`listWithLgl.docx`; the only one `edit_document()` admits today) |
| namespace declared, no MCE attribute or element | 27 |
| one root `mc:Ignorable`, projection drops only that attribute | 31 |
| … plus dropped ignorable attributes (`wp14:anchorId`/`editId`) | 1 |
| … plus dropped ignorable leaf elements (`wp14:sizeRelH`/`sizeRelV`), no WordprocessingML inside | 2 |
| `AlternateContent` (in a run: `Choice Requires="wps"` drawing vs VML `Fallback`) | 1 (`do-not-capture-draw-objs-on-page-draw-wrap-none.docx`) |

Every MCE-bearing main part declares the namespace exactly once, on the root,
and `mc:Ignorable` never appears anywhere but the root. Unmanaged
`edit_document()` refuses 60 of the 63 with the MCE `UnsafeEdit`, and 2 cannot
be opened at all. This matches change 0653's note that "both source-backed DOCX
routes refuse 61 of 62 fixtures"
(`docs/performance/0653-mce-namespace-emission-rewrite.md:743-747`). Raw and
visible transaction layouts (direct-body paragraph, table and block-control
counts) are identical on all 59 analyzable files. Paragraph text differs in
exactly one paragraph, the `AlternateContent` one (W1 below).

The fixture set under-represents modern Word output. Word 2010+ wraps
`wps` shapes and text boxes, `w16se` emoji, ink and 3D models in
`AlternateContent` inside runs, and the repository has one such file, from
LibreOffice.

## 2. Why the transaction refuses

### 2.1 The code path

- `source_backed::Package::edit_document` and `document_snapshot`
  (`crates/litchi-docx/src/source_backed.rs:1121-1145`) call
  `main_document_snapshot_with_before`. On the unmanaged branch
  (`:2064-2085`) that function computes `visible_document_xml(raw)` and
  refuses unless the result is the same allocation (`:2068-2076`). Otherwise
  it builds the transaction snapshot **from the raw bytes**
  (`Snapshot::from_shared_xml_with_source_identity`, `:2077-2082`).
- `visible_document_xml` (`crates/litchi-docx/src/parts/document_part.rs:498-517`)
  returns `raw` only for `Cow::Borrowed`.
- `process_markup_compatibility` borrows only when the literal namespace URI
  is absent (`crates/litchi-ooxml-common/src/mce/codec.rs:722-759`). Otherwise
  it re-tokenizes and returns `Cow::Owned` unconditionally (`:760-850`), even
  when the output is semantically the input. Along the way it drops MCE
  attributes (`:1563-1568`) and ignorable non-understood attributes
  (`:1569-1581`), writes `</q>` for empty tags (`close`, `:1315-1330`), and
  skips text outside the root (`:794-798`).
- The publisher re-runs the same guard (`publish_document_commit_to_stream`,
  `source_backed.rs:1500-1555`, which re-derives the current snapshot at
  `:1506-1510`). It then publishes the raw-derived candidate
  (`write_part_overlay_shared_to_stream`, `:1546-1552`).

The doc comment states the intent: "Keeping the raw bytes prevents semantic
selectors from being applied to a different XML representation than the one
eventually published" (`source_backed.rs:1124-1127`).

### 2.2 The invariants behind the guard

Change 0653 already made the codec stop re-declaring namespaces, and nothing
in that change touches this guard. The guard defends four properties:

- **I1: coordinate identity.** A `Snapshot` stores direct-body paragraph,
  table and block-control ranges as offsets into the exact bytes it will
  publish (`crates/litchi-docx/src/document/transaction.rs:935-951`; layout
  scan `:6221-6552`). `Patch::apply` is exact-source checked against those
  bytes. A snapshot built from the *visible* buffer would publish the
  projection. That drops `mc:Ignorable`, ignorable attributes and elements,
  and every non-selected branch, and rewrites every empty tag. This would
  violate ADR 0006's "Preserve is the default serialization mode"
  (`docs/adr/0006-validation-security-and-compatibility.md:8-11`). So the
  snapshot must stay raw, as it is.
- **I2: semantic-surface identity.** Readers see the projection.
  `Package::document()` builds `Document` from `visible_document_xml`
  (`source_backed.rs:726-737`), and so does the eager `DocumentPart`
  (`document_part.rs:927-942`). The editor sees raw bytes. Wherever MCE
  selects, unwraps or hides content the editor's scanners would classify, a
  selector resolved by the reader lands on a different object in the editor.
  ADR 0003 forbids this: "Ambiguous resolution is a conflict"
  (`docs/adr/0003-snapshots-edits-and-patches.md:92`), and the facade
  returns "a typed refusal instead of guessing" (`:118-119`).
- **I3: MCE conformance.** A document whose projection fails
  (`MustUnderstand` of a namespace the profile lacks, a malformed directive)
  must not become editable when it is not readable.
- **I4: ADR 0006.** "Markup-compatibility alternatives are edited only when
  their effective semantics are modeled; otherwise the operation is blocked
  before any bytes are published" (`adr/0006…:99-101`).

The raw edit itself is already lossless. Everything outside the splice is
copied, including non-selected branches, `mc:Ignorable` and the ignorable
namespace declarations (§3). The guard is not protecting preservation. It is
protecting I2 and I3, and it does so by refusing *any* document that mentions
the namespace.

### 2.3 Witnesses: where raw and visible really diverge

These are synthetic main parts run through `Snapshot::from_xml`, which is the
same scan and edit code both owners use, plus the codec. The probe edits with
`replace_paragraph_text`.

| Case | Raw direct-body paragraphs | Visible paragraphs | Raw edit | Verdict |
|---|---|---|---|---|
| W1: run-level `AlternateContent` (wps `Choice`, VML `Fallback`), text box in both | `HelloBOXBOX` | `HelloBOX` | refused `ComplexRun` | text doubles in raw; the grammar refuses, by luck |
| W2: body-level `AlternateContent` holding paragraphs | `P0`, `P2` | `P0`, `CHOICE`, `P2` | position 1 **accepted**, edits `P2`; visible after: `P0`, `CHOICE`, `NEW` | **wrong target** (I2) |
| W3: `mc:ProcessContent="x:wrap"` around a run | `WRAPPEDB` | `WRAPPEDB` | refused `ComplexContent` | safe direction |
| W4: ignorable `x:ext` holding a run, no ProcessContent | `HIDDENB` | `B` | refused `ComplexContent` | raw reads differ from visible |
| W5: `mc:MustUnderstand="x"` (not understood) | `A` | projection error `MustUnderstand("urn:x")` | **accepted** | edits what the reader refuses (I3) |
| W6: root `mc:Ignorable="x"`, `x:id` attribute, leaf `x:flag` | `A` | `A` | accepted; visible after `NEW` | safe |
| W7: body-level `AlternateContent`, MCE namespace written `…/200&#54;` | 2 | 2 (the codec never notices MCE) | `edit_document()` **admits** it today | the lexical fast path misses the escaped namespace |

On the corpus, W1's shape is the one real occurrence. In
`do-not-capture-draw-objs-on-page-draw-wrap-none.docx` paragraph 26, the raw
text is `<Text><Text>To make your document…` and the visible text is
`<Text>To make…`, and the transaction refuses it (`ComplexRun`).

Two conclusions follow. When the projection only drops directives, or drops
opaque content with no WordprocessingML inside (the 19 fixtures, W6), raw and
visible agree on everything the transaction scans, so the refusal is a false
positive. When it selects or unwraps (W1–W3), hides WordprocessingML (W4) or
fails (W5), a guard is needed, and per operation, not per document.

The scanners treat foreign content as opaque. `scan_document` classifies only
WordprocessingML `p`, `tbl`, `sdt` and `sectPr` direct body children
(`transaction.rs:6217-6219`, `:6372-6424`). `scan_text_owner` skips
non-WordprocessingML paragraph and run children (`:8301-8319`) and refuses
loose text (`:8420-8435`). Dropped *opaque* content is therefore invisible to
both views alike.

### 2.4 The two sibling guards

- **Document variables** (`source_backed.rs:2204-2216`) use the same pointer
  test on `settings.xml` with the baseline profile. They refuse 4 of the 19
  (`ComplexNumberedLists`, `Hyperlink`, `table-alignment`, `testComment`),
  whose settings parts carry a root `mc:Ignorable`. Three of them also carry
  `w14`/`w15` leaf elements (`docId`, `chartTrackingRefBased`).
- **The managed route.** With an `ExecutionContext`, `ensure_source_document_xml`
  (`source_backed.rs:2398`, called for reads at `:695` and for snapshots at
  `:2015`, `:2033`, `:2050`) refuses any MCE element, any MCE attribute and
  even a bare MCE namespace declaration (`ensure_source_mce_element`,
  `:2849-2912`). The reason given is "markup-compatibility preprocessing would
  require an unbudgeted owned payload". Under a context, `document()` (a read)
  and `edit_document()` refuse **all 19**, and 60 of 63 corpus files.
  `to_owned_package()` is refused for managed packages as well
  (`crates/litchi-opc/src/source_backed.rs:7890-7906`). M0a's
  `OpenOptions::execution` (`uo/m0a`: `crates/litchi-ops/src/options.rs:128-153`)
  therefore must not route DOCX opens into the managed constructors until
  stage 3 lands.

## 3. How the eager owner handles the same input

`litchi_docx::Package::document_snapshot` builds the transaction snapshot from
the raw main part (`Snapshot::from_shared_xml(main.blob_arc())`,
`crates/litchi-docx/src/package/package/document.rs:57-64`), with no MCE call
and no refusal. `apply_document_patch` publishes the candidate's raw-derived
bytes (`:96-144`), reusing the retained candidate when the retained main part
still matches the patch source byte for byte (`:79-86`).

The probe edited one paragraph per fixture through the eager route:
`replace_paragraph_text` on the first paragraph the grammar accepts, or
`insert_paragraph` at the end where none is accepted. It then saved and diffed
the main part and every ZIP member against the original. Results for the 19
fixtures and the 4 wider-corpus MCE shapes (`saut_page.docx`, the
`AlternateContent` file, and the two `wp14`-drop files):

- **Main part.** The changed window lies inside the edited paragraph (12
  replacements, 10 of them in the 19) or is one insertion immediately before
  the body-final `w:sectPr` (11 appends, 9 of them in the 19). The root start
  tag is byte-identical, including `xmlns:mc`/`xmlns:ve` and `mc:Ignorable`.
  The counts of `mc:` tokens, `Ignorable` and `AlternateContent` are
  unchanged. In the `AlternateContent` file (edit of paragraph 27), the
  9,139-byte `mc:AlternateContent` element in paragraph 26, both branches
  included, is untouched.
- **Owner independence.** The eager main part equals the output of the same
  edit on `Snapshot::from_xml(raw)` (23/23). With the MCE URI swapped for an
  equal-length unknown one (the "marker-stripped control" of change 0649),
  the source-backed publisher accepts the same edit and emits the identical
  main part (20/20). The 3 others are refused at publication: "document or
  write protection, or tracked revisions, is enabled",
  `crates/litchi-docx/src/source_backed/document_policy.rs:104-110`.
- **Other members.** The eager save keeps every other member's bytes,
  compressed size, CRC, method and order. The source-backed publisher
  raw-copies all of them.

So the eager owner preserves `mc` constructs **byte for byte**, not just
semantically. What it lacks is any I2 or I3 check. Packaged as `.docx` files,
W2 and W5 open, edit and save through the eager owner, while the
source-backed owner refuses both; for W5 that refusal is its own reader's
`Mce(MustUnderstand("urn:x"))`. (M0a's eager handle runs a text pass at open,
so it would refuse W5 at open, but not W2.) It also publishes edits into the
three `documentProtection_*` fixtures with `w:enforcement="1"` (comments,
forms, read-only) that the source-backed route refuses. ADR 0006 says "Office
protection is enforced by default" (`adr/0006…:145`).

## 4. Options

### 4.1 What already exists, and whether it applies to DOCX

- **Owner decision 0652 #1 → change 0653** (MCE namespace re-declaration;
  0588 designed it, 0649 priced it at 93.9% of a real-deck PPTX edit). This
  is the shared codec's output shape. It already applies to DOCX:
  `Paragraph::extensions` and friends moved to self-contained slices
  (0653 §3). It does not touch this refusal, because the guard compares
  allocations, and that is unchanged by how namespaces are emitted.
- **The XLSX equivalence proof (0603, extended under the same decision by
  0666).** `source_stream_admission` and `MceRewriteEquivalence`
  (`crates/litchi-xlsx/src/raw/worksheet/mod.rs:87-101`, `:130-283`) prove
  event by event that the rewrite is invisible to the worksheet parser. They
  admit namespace declarations, `mc:Ignorable` (validated exactly like the
  codec, `valid_ignorable_directive`, `:285-315`) and other prefixed
  attributes. They refuse `AlternateContent`, every other directive, unknown
  MCE attributes, prefixed elements and `&` in attribute values. **DOCX has no
  counterpart.** This is the model for option (b).
- **PPTX.** The opened transaction maps a shape chosen on the projected
  `Scene` back to its raw span with `mce::active_offsets`
  (`selected_raw_span_for_shape`, `crates/litchi-pptx/src/tag/shape/codec.rs:315-356`,
  `:399-423`). It rewrites the raw span and re-reads the staged slide
  (`crates/litchi-pptx/src/opened/transaction.rs:210-248`). This is the model
  for option (a). Atomic batches and source-backed slide edits still refuse
  rewritten scenes (`opened/xml.rs:781-786`; `presentation/transition.rs:36-55`).
  `has_mce_markup` (`transition.rs:57-84`) is the declared-only precedent:
  "A namespace declaration alone … cannot select a branch".
- **DOCX already uses `active_offsets`** (`mce/codec.rs:29-176`) for
  exact-source content-control snapshots
  (`crates/litchi-docx/src/content_control/snapshot.rs:412-470`, `:948-1030`),
  altChunk anchors including the source-backed altChunk publisher
  (`alt/codec.rs:97`), ink (`ink/host.rs:799`) and revision conflicts
  (`revision/conflict/codec.rs:764`). The bounded streaming processor with
  raw and active observers (`mce/stream.rs:832-930`) is available for a
  budgeted scan that does not build an output buffer.

### 4.2 The options compared

| Option | Correctness risk | Preservation | Performance (vs today's source-backed route) | Size | Unlocks (of 19) |
|---|---|---|---|---|---|
| **(a)** MCE projection: semantic positions from the visible tree, raw spans by `active_offsets`, refuse edits whose raw range meets `AlternateContent`/`Choice`/`Fallback`/unwrapped/dropped content, re-project the candidate and read back | Medium. A new coordinate layer across all 20 `Operation` kinds. Position meaning changes where body-level `AlternateContent` exists (W2). Needs a readback proof per commit. | Byte-exact outside the splice; branches untouched | Snapshot pays `active_offsets` ≈ 1.2× the full MCE pass (2.04 ms for drawing's 288 KB part, 5.6× its layout scan), and commit pays a re-projection. Transient memory ≈ marked copy + visible copy ≈ 2.2× the part | Large: est. 1.5–3k lines plus tests in `litchi-docx` | 19, but none needs it; 1 of 63 corpus files does |
| **(b)** Ignorable-only admission: one proof scan; admit when the only MCE constructs are namespace declarations and `mc:Ignorable` (validated like the codec), and everything dropped is the directive, ignorable non-understood attributes, or ignorable non-understood subtrees with no element from an understood namespace inside. Edit and publish raw bytes. | Low. The scanners already treat that content as opaque (§2.3); W6 passes; W1–W5 and W7 are refused | Byte-exact (identical to the eager output, §3) | Replaces today's full MCE pass + visible allocation with one scan without an output buffer. The prototype costs 0.30–0.37× the full pass (634 µs vs 1,714 µs on drawing); declared-only needs only a lexical check | Small: ~200–400 lines plus tests (a shared proof in `litchi-ooxml-common::mce`, three call sites) | **19** (13 at stage 1, 6 more at stage 2); publication 16 (3 protection refusals by policy) |
| **(c1)** At first edit, `to_owned_package()` and the eager save path | Medium: inherits the eager owner's missing I2/I3 checks (W2, W5) and its missing protection check (3 fixtures). Refused for managed packages | **Physical layout lost** on 23/23 outputs: member order rewritten, `[Content_Types].xml` and every `.rels` regenerated, 0–1 of N members passed through compressed, explicit directory entries dropped (`table-alignment`, `table-indent`); `saut_page.docx` grows from 2,959,626 to 3,013,162 bytes (+1.8%; the eager save gives 2,959,645). Part contents are kept | Lifecycle 0.9–6.9× the source-backed one on the four small packages (the high end re-deflates a stored 48 KB PNG), **273×** on `saut_page.docx` (74.8 ms vs 0.27 ms: every part decoded and re-deflated); `to_owned_package` alone 10.1 ms there; peak RSS +14.2 MiB vs +6.1 MiB | Tiny (unified layer only) | 19 (all published, incl. 3 protected — wrongly) |
| **(c2)** At first edit, re-open the retained source with the eager constructor (`Package::from_reader`) | Same as (c1) | Kept (eager open from bytes passes members through) | Lifecycle 0.67–0.92× the source-backed one on 4 fixtures and 1.56× on `saut_page.docx`, plus the original source open (ADR 0030 lazy decode makes the eager owner cheap). The whole file becomes resident: +2.8 MiB peak over the source-backed route for the 2.96 MB file | Small (unified layer keeps the `ReadAt`, re-checks freshness) | 19 (same caveats as c1) |
| **(d)** Stage (b): declared-only first, then Ignorable-only. Apply the same admission to the eager owner and to the managed scanner. Memoize the verdict per source revision. Add a mapped run-level `AlternateContent` stage later | Lowest per stage | Byte-exact | Stage 1 is ~free; stage 2 as (b); reads can also skip the visible copy for admitted parts | Small per stage | 13 → 19; managed 0 → 13 → 19 |

### 4.3 Measurements behind the table

Release build, pinned to CPU 21, two processes per row. Each cell is the mean
of the two process p50s in µs. The spread between the two processes is at
most 4.3% (the MCE pass on `FancyFoot.docx`), and most cells are within 2%.
"Control" is the file with the MCE URI replaced by an equal-length unknown
URI, so the current guard passes. Row E uses the original file throughout,
because the eager leg on the control copy of `saut_page.docx` was bimodal
across processes in earlier runs (≈ 385 or ≈ 1,190 µs).

| Leg | FancyFoot (2,891 B part) | testComment (3,108 B) | table-indent (71,991 B) | drawing (288,070 B) | saut_page (2,459 B part, 5.37 MB package) |
|---|---:|---:|---:|---:|---:|
| source-backed open | 41.7 | 51.6 | 33.1 | 72.2 | 88.5 |
| today: open + `edit_document` → refusal | 70.2 | 87.8 | 509.3 | 2,119.4 | 121.2 |
| control: open + `edit_document` → `Edit` | 53.3 | 62.1 | 129.6 | 584.2 | 99.4 |
| full MCE pass (today's guard) | 21.6 | 26.4 | 460.6 | 1,713.7 | 23.5 |
| (b) prototype proof scan | 7.5 | 8.4 | 169.0 | 634.3 | 7.0 |
| (a) `active_offsets` over `<w:p>` starts | 26.7 | 30.2 | 546.1 | 2,037.6 | 27.5 |
| raw layout scan (`Snapshot::from_xml`) | 5.9 | 4.6 | 85.7 | 362.3 | 4.8 |
| B: source-backed edit lifecycle (control) | 152.4 | 191.6 | 700.9 | 4,239 | 273.9 |
| E: eager edit lifecycle (open, edit, commit, publish, save) | 139.5 | 169.8 | 468.7 | 3,250 | 427.7 |
| C1: source open + `to_owned_package` + eager edit + save | 476.0 | 1,331.2 | 629.7 | 3,987 | 74,849 |
| `to_owned_package` alone | 93.3 | 116.9 | 61.5 | 351.1 | 10,103 |

Peak RSS over 20 lifecycles in one process (the baseline process holds the
input twice):

- `saut_page.docx`: baseline 8.1–8.2 MiB, B 14.0–14.4 MiB,
  E 16.9–17.0 MiB, C1 22.3–22.4 MiB;
- `drawing.docx`: 6.7–7.0 MiB for all three.

Observations, not claims:

- Today a caller pays a full MCE pass and a discarded visible copy just to
  receive the refusal (2.12 ms vs 0.58 ms on drawing).
- At leaf level the eager lifecycle is not more expensive than the
  source-backed one here. The source-backed publisher re-derives the snapshot
  (`source_backed.rs:1506-1511`), a second layout scan plus the settings policy
  scan, where the eager route reuses its candidate. The facade-level gap
  measured in S1 (`docx_file_eager_open` 11,498 µs vs `docx_file_source_open`
  316 µs on `docx-source-backed-media`, `uo/m0a`:
  `docs/unified-ops/results/0003/before/README.md`) includes the eager route's
  validation text pass.

### 4.4 Notes on each option

**(a)** This is the general solution and the right end state for
`AlternateContent`. It is not the right first step: no fixture needs it, its
coordinate layer has to cover every operation, and W2 shows that "position"
must be defined over the visible tree. That definition changes the meaning of
eager positions for documents with body-level `AlternateContent`. A restricted
form, stage 4 below, covers the common Word case (`AlternateContent` inside
runs) without remapping positions, because those documents keep identical
body layouts (59/59 corpus files, including the `AlternateContent` one).

**(b)** The proof must reproduce the codec's directive validation:
NCName tokens, no duplicates, bound prefixes, never the MCE namespace, and the
token limits. It must refuse everything else the codec would act on:
`AlternateContent`/`Choice`/`Fallback`, `ProcessContent`, `MustUnderstand`,
`PreserveElements`/`PreserveAttributes`, and unknown MCE attributes. It must
decode namespace declarations before comparing (W7), as the managed scanner
already does (`source_namespace_binding_is_mce`, `source_backed.rs:2922-2950`).
Restrict it to an `mc:Ignorable` on the document root, which covers 35 of 35
corpus occurrences and avoids scope questions when content moves. The
"no understood descendant in a dropped subtree" clause excludes W4. That clause
is what makes reads over raw bytes equal reads over the projection, which also
lets the managed route read admitted parts without the projection.

**(c)** A fallback cannot be safer than the admission it bypasses. Once (b)
exists, both owners need the same admission, and the fallback adds only cost:
(c1) loses ZIP framing and pays to decode everything, (c2) holds the whole
file. Silently switching owners also changes the protection verdict for 3
fixtures. Note that M0a's `SourceMode::Eager` path already uses
`to_owned_package()` (`uo/m0a`: `crates/litchi-ops/src/format/docx.rs`,
`Docx::from_source`), so an eager handle opened that way saves with
regenerated framing (finding A5).

## 5. Recommendation and staged plan

Adopt (d): stage (b), make it the single admission for both owners, and keep
`AlternateContent` refused until stage 4. Do not add an eager fallback. A
document refused by the admission keeps its typed `UnsafeEdit`, which maps to
`unsupported_capability` with a suggestion. Each stage is a leaf change in
`litchi-docx` (with the reusable proof in `litchi-ooxml-common::mce`), as the
goal prefers (`docs/UNIFIED_OPS_GOAL.md:105-107`).

**Stage 1: declared-only admission.** If the decoded MCE namespace is
declared but no element or attribute is in it, the projection is lexical only,
so admit raw bytes. Apply this at the unmanaged guard
(`source_backed.rs:2068-2076`), the document-variables guard (`:2209-2216`),
and the managed scanner, which should stop refusing a bare declaration
(`:2884-2909`) but keep refusing MCE elements and attributes.

- Unlocks **13**: `FancyFoot`, `Headers`, `documentProperties`,
  `documentProtection_comments_no_password`,
  `documentProtection_forms_no_password`,
  `documentProtection_no_protection`,
  `documentProtection_readonly_no_password`, `drawing`, `endnotes`,
  `footnotes`, `headerFooter`, `headerPic`, `table_footnotes`. This holds for
  `edit_document()` and, in managed mode, for `document()` too. Publication
  succeeds on 10. The 3 protected fixtures keep their typed protection
  refusal.
- Corpus: +25. That is all 27 declared-only files except the 2 that cannot
  be opened, so 26 of the 61 openable files are admitted, counting the
  MCE-free one.

**Stage 2: Ignorable-only proof.** Admit a root `mc:Ignorable` whose
projection drops only the directive, ignorable non-understood attributes, and
ignorable non-understood subtrees without understood descendants (§4.4).

- Unlocks the remaining **6**: `ComplexNumberedLists`, `Hyperlink`,
  `comment`, `table-alignment`, `table-indent`, `testComment`. That makes
  19/19 for `edit_document()` and 16/19 for publication.
- Document variables become 19/19. The 4 refused today need the
  opaque-leaf clause, because the baseline settings profile drops their
  `w14`/`w15` leaves.
- Corpus: +34, so 60 of the 61 openable files, every one except the
  `AlternateContent` file.

**Stage 3: one policy for both owners and for the managed route.**

- Run the same admission in the eager owner's `edit_document()`, so the eager
  owner refuses W2, W5 and W7 shapes as the source owner does. This is a
  behaviour change for eager users of `AlternateContent` documents, accepted
  under the owner's "correctness first" trade-off (record 0652).
- Implement the stage-2 proof inside `ensure_source_document_xml` under the
  existing budget charge. Managed reads and edits then reach 19/19 (13 after
  stage 1, 0 today).
- In the unified layer, apply the publication protection policy whichever
  owner holds the document.
- No new fixtures. Changes the verdict on the 3 protected fixtures for the
  eager owner, and takes managed reads from 13 to 19.

**Stage 4: run-level `AlternateContent`.** Admit documents whose every
`AlternateContent` lies inside a `w:r`. Refuse any operation whose changed
byte ranges intersect an MCE element. After commit, re-project the candidate
and require the expected visible effect. This is the PPTX `set_shape_text`
pattern.

- 0 of 19; 1 corpus file. It needs new Word-authored fixtures (text box,
  `wps` shape, `w16se` emoji, ink) before it starts.
- Body-level `AlternateContent`, `ProcessContent` and unsatisfiable
  `MustUnderstand` stay refused. That is (a)'s territory, and the corpus has
  none.

**Tests that would prove each stage.**

- *Admission matrix (unit).* Admitted: declared-only (`ve` shape; `mc` bound
  but unused), W6, the wp14 attribute and leaf-element drops. Refused with
  today's error identity: W1, W2, W3, W4, W5, W7, `Preserve*`, unknown MCE
  attribute, Ignorable listing an unbound prefix, duplicate prefixes, the MCE
  namespace itself, Ignorable on a non-root element.
- *Corpus differential (the 0603/0666 pattern).* Over every DOCX-family
  fixture:
  - admission verdicts;
  - for admitted files, source-backed edit → commit → publish emits the same
    main part as the eager route and raw-copies every other member;
  - exact no-op republishes the source byte for byte;
  - the 3 protected fixtures keep the identical typed refusal.
- *Projection oracle (tests and debug builds).* For every admitted part,
  `Snapshot::from_xml(raw)` and `Snapshot::from_xml(visible)` agree on layout
  and paragraph text. This mirrors PPTX's debug cross-check
  (`tag/shape/codec.rs:326-337`).
- *Managed.* Budget charges for the proof scan, cancellation, and refusal
  identity for non-admitted parts. Extend
  `crates/litchi-docx/tests/source_backed_managed_document_edit.rs`.
- *Unified layer.* Invert M0a's `source_editor_refusal_on_markup_compatibility_is_typed`
  (`uo/m0a` `tests.rs:400-429`) to expect success on its 4 fixtures, and keep
  a refusal test on a synthetic `AlternateContent` document.
- *Performance.* A paired leaf measurement of `edit_document` before and
  after on these fixtures. The bar: the admission scan costs no more than the
  layout scan it accompanies, and no fixture is slower than its current
  refusal path.

## 6. Adjacent findings for M1/M2

- **A1.** The managed source-backed route cannot read any MCE-bearing Word
  file (60/63 corpus files) (§2.4). Keep DOCX on the unmanaged constructors
  until stage 3.
- **A2.** The eager owner publishes edits into documents with enforced
  protection; the source-backed owner refuses them. The unified layer needs
  one policy (ADR 0006:145).
- **A3.** The codec's fast path is lexical (`mce/codec.rs:751`), so an
  entity-escaped MCE namespace disables MCE processing for every reader (W7).
  This needs a shared-codec follow-up, which moves behaviour for such inputs.
- **A4.** Two paragraph index spaces exist. The source-backed `Document`
  counts outermost paragraphs anywhere, table cells included
  (`crates/litchi-docx/src/namespace.rs:432-655`, via
  `parts/document_part.rs:520`). The transaction counts
  direct-body paragraphs (`transaction.rs:6372-6424`). M1 must bind
  `/body/p[N]` to one of them for both `get` and `set`.
- **A5.** `to_owned_package()` loses ZIP framing (§4.2). M0a's
  `SourceMode::Eager` route uses it, so saving such a handle would violate
  "compressed-entry passthrough for untouched members"
  (`UNIFIED_OPS_GOAL.md:142-143`).
- **A6.** The source-backed publisher rescans the main part that the eager
  route reuses (§4.3). This is a possible saving; no claim is made.

## 7. Evidence and reproduction

Nothing below is committed. The probe, its outputs and the build directory
were deleted after this record was written.

- **Probe source.** `crates/litchi-docx/examples/uo0008_mce_probe.rs`,
  throwaway, 670 lines, SHA-256 `a7d0bb37da651a446bf029153b22f9072c09bc482461b1c75db6fc0191ee9a23`.
  Built with:

  ```sh
  CARGO_TARGET_DIR=$PWD/target CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 \
    cargo build --release -p litchi-docx --example uo0008_mce_probe --offline
  ```

  The binary's SHA-256 was `fc1cefab1b88b1c2357c48c14255fb8f9aceca006178f810e66d86c70cb3dfcb`.
  The toolchain was rustc 1.95.0 on an AMD EPYC 9R45 host running Linux
  7.0.0-1013-aws. Load average was 0.8–3.6 during measurement.
- **Probe modes.** All use only public APIs:
  - `census <out> <docx…>`: the §1 and §3 results;
  - `witness`: W1–W6;
  - `classify <docx…>`: the prototype proof scan and its costs;
  - `time <orig> <control> <paragraph> 400`: the §4.3 legs, warm-up 3, up to
    3,000 samples or 400 ms;
  - `rss <route> …`: under `/usr/bin/time -f %M`, 20 lifecycles.
  W7, and the packaged W2 and W5 checks in §3, were `census` runs on
  minimal synthetic packages.
- **Python helpers** (lxml 6.0.2):
  - the classifier (SHA-256 `d27d62d0…c9db`) for the construct counts;
  - the control maker (`f02e154f…dc48`), which replaces the MCE URI in the
    main part only;
  - the differ (`5ee97922…3023`) for the member and window checks;
  - an inline check for the lexical-rewrite equivalence in §1.
- **Re-creating the probe.** Each mode is a few public calls:
  `source_backed::Package::from_read_at[_with_execution_context]`,
  `edit_document`, `document`, `document_variables_snapshot`,
  `publish_document_commit_to_stream`, `to_owned_package`;
  `Package::from_reader`, `document_snapshot`, `edit_document`,
  `publish_document_commit`, `to_stream`; `Snapshot::from_xml`,
  `Edit::{replace_paragraph_text, insert_paragraph, commit}`;
  `mce::{process_markup_compatibility, active_offsets}`.

## What is not claimed

- No timing, allocation or RSS number is a registered performance claim.
  They come from one host, with two process repeats, a warm cache and
  in-memory sources.
- The prototype proof scan is unoptimized and was not checked against the
  codec beyond the 23 files and 6 witnesses. Its cost is an estimate for
  option (b), not a measurement of a design.
- The implementation sizes are estimates.
- No Microsoft Word round trip of any probe output was performed.
- The 63-file corpus is not representative of modern Word output for
  `AlternateContent` (§1).
