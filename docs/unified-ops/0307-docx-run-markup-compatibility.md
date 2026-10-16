# 0307 — DOCX run-level markup compatibility

- Status: implemented on `uo2/docx-mc`, cut from `feat/unified-ops-api` at
  `e7fe30fb24`, trunk merged at `208e74e6ce` (record 0303). Follows record
  [0303](0303-pptx-docx-producer-structure-gaps.md) §5, record
  [0293](0293-producer-structure.md) gap 5 and record
  [0285](0285-find-merge-consolidation.md) §8.
- Date: 2026-10-16
- Owned: the `litchi-docx` document-transaction admission, position verdict
  and scanners; one additive proof in `litchi-ooxml-common::mce::admission`;
  the DOCX find adapter and capability-gap texts in `litchi-ops`.
- `performance_claim`: none. §6 shows no DOCX set/add path got slower beyond
  noise.

## 1. Result

Word writes drawings and text boxes as an `mc:AlternateContent` inside a
run: a DrawingML `Choice` and a VML `Fallback`, each with its own copy of
any text-box text. Document transactions refused every such document, so
the unified API could not edit it at all. They are now admitted, and each
wrapper is one opaque inline object.

| Operation on a document with run-level alternate content | Before | After |
|---|---|---|
| any edit (`edit()`) | refused: `unsafe DOCX edit … markup-compatibility branch selection` | admitted |
| body `set`/`remove`/`move`, paragraph and run property sets | refused at admission (raw positions not proven) | work; every wrapper is kept |
| body `add` (paragraph, run beside a wrapper run) | refused | works; never splits a wrapper run |
| `text` set on a paragraph or run whose wrapper holds no text | refused | works; the wrapper is kept |
| `text` set over a wrapper whose branches hold text (a text box) | refused | refused, typed: `Refusal::AlternateContentText` (`unsupported_capability`) |
| find/replace | refused | works; wrapper text is not searched; a match that would span a wrapper is refused |
| template merge | refused | works; placeholders inside wrappers are left as they are |

Corpus: all five `test-data` DOCX files with run-level wrappers (§4) are
admitted. The fidelity report's DOCX refusals drop from 1 to 0.

## 2. The three proofs

### 2.1 Admission

`litchi_ooxml_common::mce::admission::admit_held(xml, capabilities, limits,
holder)` repeats the shared admission proof (record 0011) with one more
shape. An `mc:AlternateContent` that has a `Holder` ancestor and lies
outside any subtree processing drops is opaque:
- the declarations on its own start tag are read;
- nothing inside it is examined beyond well-formedness, the depth bound and
  the visible-reference check.

Every other markup-compatibility element and directive is refused as
`admit` refuses it. The result, `Held { outside, wrappers }`, gives the
`admit` verdict outside the wrappers and the count of outermost wrappers.

The proof does not establish that processing succeeds inside a wrapper.
For example, a `Requires` that names an unbound prefix is a processing
error. The caller runs the processor.

`litchi-docx`'s `admit_document_source` keeps its order:
1. the plain proof;
2. otherwise the projection, which returns its error as is;
3. a projection identical to the source is admitted.

A changed projection used to be refused. It is now admitted as
`SourceAdmission::RunAlternateContent` when `admit_held` with the run
holder (`w:r`, Transitional or Strict) finds at least one wrapper. The
same rule serves both owners: the eager `Package::document_snapshot` and
the source-backed snapshot.

Two cases stay refused:
- **Managed (execution-context) documents.** They have no projection, so
  they are still admitted only when the projection would differ in
  spelling alone.
- **Paragraph transfer from such a donor** (`TransferRefusal::
  MarkupCompatibilityScope`). The donor's selection is not a lexical
  property of the copied bytes.

### 2.2 Position verdict

`VisiblePayload::project` marked raw positions `Unproven` whenever the
projection selected a branch. A selected branch replaces its wrapper inside
the run, so every block and every run keeps its parent and its order. The
verdict is now `Preserved` when the projection selected branches and
`admit_held` proves every wrapper run-level. This is the same rule as
admission, so "admitted ⇒ positions proven" still holds.

The corpus invariant in `tests/body_block_index.rs` compares each raw
paragraph's text with the visible paragraph's text. For a paragraph holding
a wrapper, the raw reader reads both branches and the visible one reads
only the selected branch, so the two texts are now compared outside inline
objects (drawings, pictures, embedded objects, wrappers). Block positions,
kinds and counts are compared exactly as before. On the corpus: 75 files,
75 admitted, 12 such paragraph checks (six paragraphs, both owners).

### 2.3 Opaque runs

- **Story scanner** (`document::substitute`, used by find and merge). A
  wrapper anywhere is skipped to its end tag:
  - a wrapper that is a direct run child reads as one `OBJECT`;
  - its offset is listed in `Paragraph::alternate_content`;
  - the text-box paragraphs of either branch are not listed.

  This applies to main and part stories (headers, footers, notes,
  comments).
- **Text-owner scanner** (paragraph, run, hyperlink, control and cell text
  sets):
  - a wrapper is skipped, so the drawing's numeric text such as
    `wp:posOffset` no longer refuses the set as `ComplexRun`;
  - a wrapper holding a WordprocessingML text element (`t`, `delText`,
    `tab`, `ptab`, `br`, `cr`, `noBreakHyphen`, `softHyphen`, `sym`) in
    either branch refuses with `Refusal::AlternateContentText`. Readers
    show the selected branch's text as part of the paragraph's, so
    replacing the paragraph's text could only keep that text, failing the
    readback, or drop the object.

  A fragment is parsed without its inherited declarations, so an unbound
  prefix on `AlternateContent` counts as markup compatibility, and a
  prefix bound to another namespace does not.
- **Structure edits** need no change. The layout scanner indexes only
  body-level blocks, and run insertion places a whole run between direct
  runs, so neither can split a wrapper run.
- **Run text set on the wrapper's own run.** It is refused, through the
  text-owner refusal above, or by `literal_run_direct_child_refused`
  before it.

### 2.4 Find spans (ops)

The leaf ends a text group at every `OBJECT`, as at a picture, so a find
never plans a match across a wrapper, and a replacement never splits or
drops one. Readers, though, show the wrapper's selected branch inline, or
nothing for a shape without text. The DOCX find adapter
(`refuse_alternate_spans`) therefore matches each paragraph that holds a
wrapper once more, with the groups joined across wrappers. If any match
crosses a joint, the find is refused (`unsupported_capability`, property
`find`) instead of being reported as no match. The check is conservative:
for a text box it can refuse a needle the visible text does not contain.

### 2.5 Compaction fix found on the way

The paragraph compactor, which runs on every changed paragraph,
re-escaped attribute values without their whitespace references. A tab,
line feed or carriage return written as `&#x9;`, `&#xA;` or `&#xD;` was
emitted literally, and attribute normalization reads a literal one as a
space. VML `o:gfxdata` base64 in a fallback branch carries such line
breaks, so a text set beside a text box changed the fallback's data. They
are now written back as references (`writer::doc::codec`).

## 3. API changes

| Item | Kind |
|---|---|
| `litchi_ooxml_common::mce::admission::{admit_held, Holder, Held}` | new, additive |
| `litchi_docx::document::substitute::Paragraph::alternate_content` | new, additive |
| `litchi_docx::document::Refusal::AlternateContentText` | new variant (the enum is `non_exhaustive`) |
| `litchi-docx`: crate-private `SourceAdmission`, `holds_run_alternate_content` | internal |

Glossary rows are added in [API_GLOSSARY](../API_GLOSSARY.md) under
"DOCX run-level markup compatibility". The capability gap texts
`set paragraph wrapped-or-unproven-selection` and
`set paragraph/run complex-leaf-content` now describe the run-level case.
The goldens are re-blessed in both feature sets, and
`officecli-capability-delta.md` is regenerated (reason text only).

## 4. Tests and fixtures

Fixtures with run-level wrappers in `test-data/libreoffice-core/sw/qa/…`
(MPL-2.0; provenance with SHA-256 in `test-data/libreoffice-core/PROVENANCE.md`):

| Fixture | Content | Status |
|---|---|---|
| `core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx` | text box, `<Text>` in both branches | already tracked |
| `extras/ooxmlexport/data/tdf90153.docx` | text run + shape without text | copied |
| `extras/layout/data/inline-textbox-position.docx` | two text boxes in one paragraph | copied |
| `writerfilter/dmapper/data/redlined-shape-sdt.docx` | group inside `w:ins`, check-box control, tracked revisions on | copied |
| `extras/ooxmlexport/data/tdf103389.docx` | two drawing canvases, one inside a bookmark | copied |

In the untracked upstream checkouts, 336 DOCX files hold run-level
wrappers (728 wrappers, 342 with text). Two more files also hold wrappers
nested in `wp:anchor`, which are still inside runs, and one
(`HelloO14.docx`) holds a body-level wrapper, which stays refused.

The tests:
- `litchi-ops` `tests/docx_markup_compatibility.rs`. Each case runs on both
  owners and is checked after a save and reopen:
  - body set/remove/add/move/align keep every wrapper; wrappers are
    compared without inter-tag whitespace, since a changed paragraph is
    published compacted;
  - run adds at positions 0, 1 and 2 beside the wrapper run;
  - a run property set on the wrapper run, a text set on its sibling, and
    an explicit run removal;
  - a text set keeps a wrapper without text;
  - text sets are refused, untouched, on both text-box fixtures;
  - find never matches text-box text (an exact no-op) and replaces beside
    it in three scopes;
  - a formatted find inside a wrapper run splits text and keeps the
    wrapper whole;
  - spanning finds are refused;
  - a merge fills a placeholder beside a text box and leaves the two inside
    it;
  - canvas and two-text-box fixtures take add and remove;
  - the redlined fixture reaches the tracked-revision publication refusal,
    not the markup-compatibility one.
- `litchi-ooxml-common`: witness and processing tests for `admit_held`, and
  the differential (`admit_held` with no wrapper equals `admit`, and admits
  a wrapper only where `admit` refuses an element).
- `litchi-docx`:
  - position-verdict, story-scanner, text-owner, compactor and transfer
    tests;
  - `mce_admission`'s W1 witness moves from refused to admitted, and a
    paragraph-level (outside a run) witness stays refused;
  - the `body_block_index` corpus invariant (§2.2).

Re-blessed after review:
- the corpus fidelity report: DOCX exact 69 → 74 and refused 1 → 0
  (`do-not-capture…` is byte-exact, and the four copies are new rows);
  revert byte-exact 36 → 39; one new `refused` revert for the redlined
  fixture's tracked revisions;
- the DOCX rows of `tests/golden/query_corpus/digests.json`. Only the four
  new fixtures change them: with the fixtures moved aside, the old digests
  pass unchanged (release run, 26.8 s).

The CRUD checklist is unchanged, since its recipes run on documents Litchi
made.

## 5. Coordination with record 0304

Record 0304 (text-set content loss) is not on trunk yet. This record
changes only the markup-compatibility layer of the text sets: the
wrapper skip and the `AlternateContentText` refusal inside
`scan_text_owner`. It does not touch the cell-text paths 0304 rewrote, so
when 0304 lands the two should compose without conflict. 0304's test file
asserts that pictures in runs survive text sets, which matches §2.3. On
merge, re-bless the fidelity report again: both records change DOCX rows.

## 6. Measurements

`crates/litchi-ops/examples/docx_mc_bench.rs`: release, `taskset -c 9`,
nine samples of 20 calls after two warm-ups, with leaf and unified legs
alternating ABBA inside each binary. Twelve runs alternated
before/after/after/before. "before" is the base `e7fe30fb24` plus the
example; "after" is `9c0f50308f` (the code commits). Evidence:
[`results/0307/summary.txt`](results/0307/summary.txt) and its twelve
JSON runs. Medians in µs per call, over the six runs of each binary:

| Scenario (1,000-paragraph document) | Leg | Before | After | Δ |
|---|---|---|---|---|
| text set, middle paragraph | leaf | 768.5 | 765.9 | −0.3 % |
| | unified | 1918.6 | 1897.7 | −1.1 % |
| paragraph add, middle | leaf | 1120.1 | 1114.6 | −0.5 % |
| | unified | 2365.3 | 2337.4 | −1.2 % |
| run `bold` set (compacts the paragraph) | unified | 1538.9 | 1524.2 | −1.0 % |
| find/replace over the body | unified | 5297.2 | 5290.3 | −0.1 % |
| text set, `do-not-capture…` `/body/p[28]` | unified | refused | 293.3 | — |
| paragraph add at start, same fixture | unified | refused | 251.4 | — |

No regression: every delta is noise-sized and favours the new build.

The unified/leaf ratio on this 1,000-paragraph text set (2.48) and add
(2.10) is the same before and after. It is outside the goal's 10 % small-edit
budget, but it is neither introduced nor changed here. Documents without
alternate content run the admission they ran before. A document with
wrappers additionally pays one `admit_held` pass in admission (memoized
per source version) and one in the position verdict (memoized per payload
generation). Binaries are not committed. SHA-256: before
`d8ca160d51fe7ada971368616cdafe0ece32306cce08ea446ebccd82633931fe`, after
`9579372950a9ced01bbb59c95cea39803b7c31698c9e5cc29d6d74622cb1cda8`. Build:
`cargo build --release -p litchi-ops --features docx --example
docx_mc_bench`.

## 7. Gates

All exit 0 on the final tree (after the trunk merge), with
`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8` and one cargo command at a
time.

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ooxml-common -p litchi-docx -p litchi-cli -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features docx \| every feature but encryption] --all-targets -- -D warnings` | 0 each |
| `cargo test -p litchi-ooxml-common -p litchi-docx --all-features --no-fail-fast` | 0 |
| `cargo test -p litchi-ops --all-features` (`--lib`, `--doc`, then the 135 integration tests in chunks of 15) | 0 (before the merge, `query_corpus_differential` failed on the four new fixtures and was re-blessed, §4) |
| `cargo test -p litchi --all-features --no-fail-fast` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ooxml-common -p litchi-docx -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `python3 tools/test_no_3rdparty_references.py`; `python3 tools/unified_bench_matrix.py --check` | 0 each |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (279 tests) |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry`, `--all-features` and every feature but `encryption`; `tools/officecli_capability_delta.py` | 0; DOCX reason text only (§3), clean on rerun |
| `LITCHI_BLESS=1 … --test corpus_fidelity --test crud_matrix` | reviewed (§4) |
| `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus`; `python3 tools/officecli_corpus.py`; `--check` | 0, 0, 0 (unchanged) |

## 8. Gaps and follow-ups

- **Text-box stories.** The text inside a wrapper is not editable. Doing
  that needs both branches edited consistently: Word reads the `Choice`
  and older readers the `Fallback`. Follow-up: "0307 text-box stories".
- **Text sets over text boxes** are refused (typed) rather than keeping
  the box and its text. A keep would break the semantic readback, and so
  break `get` after `set`, because readers count the box's text.
- **Paragraph-level and body-level wrappers** (`mc:AlternateContent` as a
  `w:p` or `w:body` child) stay refused, as do `ProcessContent`, nested
  foreign `Ignorable` and `MustUnderstand` outside wrappers.
- **Managed (execution-context) documents** with wrappers stay refused at
  read: there is no budgeted projection.
- **Paragraph transfer** from a donor with wrappers is refused.
- **`Snapshot::paragraph(..).text()`** on a raw paragraph that holds a
  wrapper reads both branches. That reader view is the raw one, and edits
  never use it for wrapper text.
- **Find span refusal is conservative** (§2.4).
- **Non-MC drawings.** A text set over a paragraph whose drawing is not
  wrapped still refuses as `ComplexRun` when the drawing holds character
  data, as before.

## 9. Not claimed

- Behaviour on run-level wrappers in headers, footers and notes is covered
  by the story scanner (find and merge skip them), but no part-story edit
  other than find and merge was exercised on such parts.
- The 336-file upstream census was not run through the unified API; only
  the five tracked fixtures were.
