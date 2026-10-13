# 0276 — Unified ODT adapter: body index, reads and writes

- Status: implemented on `uo2/s23-odt`, cut from `feat/unified-ops-api` at
  `395d0ffa9f`. Plan [0248](0248-coordinator-restart-plan.md) §6 Phase F,
  step 23 (ODT stream); builds on the ODF W3 package split of
  [0268](0268-odf-w3-package-split.md).
- Date: 2026-10-13
- Scope: `crates/litchi-odt` (new `body` and `metadata` modules, body edits on
  the primary `Edit`, `SourceBackedDocument::to_package`),
  `crates/litchi-ops/src/format/odt*`, and small separate hunks in the shared
  dispatch files (`read`, `write`, `transaction`, `save`, `create`, `open`,
  `open/path`, `capabilities`, `error`, `node`, `node/view`, `query/{matches,
  plan}`, `view/json`, `path/vocabulary`). Tests: the ODT unit and
  integration suites, the query-corpus differential, capability and registry
  goldens, one ODT round trip in the CLI (`litchi-cli/tests/cli.rs`) and the
  Python suite (`pyo3-litchi/tests/test_write_api.py`). No iWork crate, no
  `litchi-odf-common` change, no sibling ODF adapter file.
- `performance_claim`: paired leaf/unified medians in §6 (release, one idle
  pinned core, ABBA ×9 with 2 warm-ups, three runs). No frozen-harness
  qualification.

## 1. What landed

Before this record ODT opened and read text, counts and metadata only;
`get`/`query`/`view` were refused (“M1 R7 adapter has not landed”) and
`capabilities()` was a whole-format refusal. ODT now has unified read and
write on the P0 surface:

| Verb | ODT support |
|---|---|
| `get`, `find`, `query`, `children`, node views | every kind of §2 |
| `view` | `text`, `annotated`, `outline`, `stats` (JSON and text) |
| `set` | text, paragraph and character formatting, styles, heading levels, hyperlink targets, names, section flags, frame geometry, `/metadata` |
| `add` | paragraph, heading, run, hyperlink, bookmark, table, row, list, list item, section, footnote, endnote, comment |
| `remove`, `move`, `swap` | every body element except cells (grid) and rows with vertical merges |
| `edit`, `batch` | atomic; direct-path sets coalesce into one leaf commit |
| `save`, `write_to`, `to_bytes` | the leaf package bytes; atomic save, `Overwrite::Refuse` honoured |
| `create` | empty text document, with create-time metadata |

## 2. The leaf body index (`litchi_odt::body`)

`Tree` indexes `office:body/office:text` of `content.xml` in one quick-xml
pass, plus the automatic styles of `content.xml` and the common styles of
`styles.xml`. Each node keeps its kind, parent/child links and four source
offsets (start, end of start tag, start of end tag, end); text, attributes
and style properties are read from the source bytes on demand. The tree is
cached once per `Document` (`Document::body()`), so clones share it.

Kinds: body, paragraph (`text:p`), heading (`text:h`), span, text (character
data with `text:s`/`text:tab`/`text:line-break` directly in a paragraph,
heading or hyperlink), hyperlink (`text:a`), list, list item/header, table,
row (through header-row and row groups), cell, covered cell, section, frame,
image, text box, footnote, endnote (by `text:note-class`), annotation,
bookmark (`text:bookmark`, `text:bookmark-start`). Nested spans belong to the
outer run; notes, bookmarks, frames and hyperlinks inside a span attach to
the span's parent. Tracked changes, forms, declarations, note citations,
embedded objects and `dc:*` metadata are skipped. Visible text follows ODF
white-space processing and leaves notes, annotations, frames and tracked
deletions out of a paragraph's text (as the leaf `Document::text` does).

Namespace prefixes are read from each part's root declarations, so aliased
prefixes (0268 §4) classify correctly; attributes are scanned linearly with
no quick-xml attribute map (record 0770).

## 3. Lossless body edits (primary `Edit`)

`Edit::{set_text, set_attribute, set_text_properties,
set_paragraph_properties, insert, remove, relocate, swap}` name nodes of the
source package's tree (`NodeId`). Every edit is validated at staging against
that tree; an edit whose bytes meet another node's pending replacement is
refused immediately (claims are indexed by start offset, so staging is
O(log n)). Edits of one node merge; insertions at one offset keep sequential
`Start`/`After` versus `End`/`Before` order. Formatting derives one automatic
style per distinct (base style, edit) — copied and rebuilt compactly from an
automatic base, or parented on a common one — and a character-data run that
gains formatting becomes a `text:span`. Range bookmarks and comments lose
their end markers with them.

At commit all body edits compile to non-overlapping splices of the exact
`content.xml` (plus `meta.xml` for metadata) and publish through
`litchi_odf_common`'s `XmlSplicePublication` and
`rebuild_package_with_xml_splices`: every byte outside the touched ranges
keeps its source spelling, untouched members are copied, and an edit that
changes nothing returns the source package (exact no-op). Body edits must
precede other operations of one `Edit`, are refused by `JoinedEdit` and the
durable/transfer vocabularies (like annotation operations, 0268 G7), and one
`Edit` admits 65,536 of them (they share one publication; the other
operations keep the 1,024 ceiling).

`litchi_odt::metadata` is the ADR 0034 §7 family for `meta.xml`
(`Package::{edit_metadata, apply_metadata}`, `metadata::{Edit, Patch, Commit,
Diagnostics, Error, LimitKind}`, `Field`); it never writes dates or the
generator, so no clock is consulted. The family is registered in the `odt`
glossary profile and the glossary row flipped.

## 4. The unified adapter (`litchi_ops::format::odt`)

- **Vocabulary.** A dedicated `ODT` vocabulary extends the Word family with
  `h`/`heading`, `list`, `listItem`/`li`, `frame`, `metadata`, and the
  aliases `annotation` → `comment` and `span` → `r`; ODT paragraphs carry no
  OOXML identity. Root children are `/body`, `/styles` (common styles by
  name, `/styles/<name>`) and `/metadata`.
- **Nodes.** Positional addresses are built while navigating (no
  per-document cache needed for paths); paths of up to four steps render
  inline without allocation. Ordinal lookups past 32 siblings use a
  per-handle child index built once, so a batch addressing `/body/p[N]` for
  every N stays linear.
- **Properties** (local, as DOCX reports them): paragraph and heading text,
  style, align, spacing, indents, line spacing, keep-next, page-break-before,
  shading and the character facets of their style; run text, style and
  bold/italic/underline/strike/size/color/font/highlight/super/sub/caps/
  smallcaps; hyperlink url/tooltip/tgtFrame; table name/rows/cols/style; row
  cells; cell text/colspan/rowspan/covered/valueType/shading/valign; section
  name/protected/hidden; frame name/geometry/anchor; picture src; text box,
  note (+citation, class) and comment (+author, date) text; bookmark name;
  style name/displayName/family/parent and their properties; metadata
  title/subject/description/author/lastModifiedBy/language/keywords (+
  read-only created/modified/application/revisionNumber).
- **Writes.** Each verb resolves its targets against one handle and stages
  the translated leaf edits on one leaf `Edit`; refusals from the leaf's
  staging are `invalid_value` (`caused_by_argument`), and unknown or
  read-only keys are `unsupported_property`/`unsupported_capability`.
- **Batches.** `File::edit`/`batch` coalesce consecutive direct-path sets
  into one leaf commit. Every plan resolves against the run's base; a set
  inside content an earlier set replaced (or replacing content an earlier
  set edits) closes the run first, so later positional paths resolve against
  the document their sequential predecessors produced. Sets that would shift
  later paths (clearing a character-data run, replacing a run that holds
  notes or bookmarks) and `/metadata` run on their own.
- **Owners.** The eager route keeps the leaf `Package` and its `Document`;
  the filesystem path route keeps the `SourceBackedDocument`, builds the body
  index from its retained `content.xml`, and materializes a package once, on
  its first edit or publication (new leaf `SourceBackedDocument::to_package`,
  source-version checked). Path opens record their origin, so `save()`
  writes back.

## 5. Goldens and generated tables

- Query-corpus differential: 34 ODT selectors over the 23 `.odt` files of
  `test-data` (22 opened, 1 refused by the leaf open), recorded from the read
  build (commit `00e19ba5b5`) before the write adapter; byte-identical on
  every later run.
- Capability and registry goldens `odt.json` added; identical under
  `--features odt` and `--all-features`. No other format's golden changed.
- `officecli-capability-delta.md` regenerated: the tool already compares ODT
  against the DOCX help schema; ODT-only elements (`heading`, `list`,
  `listItem`, `frame`, `metadata`) appear as `litchi-only`, and the ODT
  contextual operation gaps are listed.

## 6. Measurements

`crates/litchi-ops/examples/odt_paired_bench.rs` (also
`tools/perf-baseline/unified-odt-paired`), release, `taskset -c 9` (an idle
core whose SMT sibling was idle; core 7 was shared with other agents and its
runs were discarded), documents of 100/1,000/10,000 builder paragraphs, ABBA
×9 with two warm-ups, medians of three runs (`results/0276/`). The harness
asserts outside the clock that the unified 1k batch publishes the same bytes
as the leaf batch.

| Scenario | Paragraphs | Leaf (µs) | Unified (µs) | Ratio |
|---|---:|---:|---:|---:|
| open + first get | 100 | 39.8 | 40.7 | 1.022 |
| open + first get | 1,000 | 137.5 | 137.0 | 0.997 |
| open + first get | 10,000 | 1130.7 | 1129.7 | 0.999 |
| full query (`p`, text + path) | 100 | 2.9 | 18.2 | 6.195 |
| full query | 1,000 | 29.9 | 178.6 | 5.962 |
| full query | 10,000 | 319.1 | 1867.3 | 5.852 |
| single set (paragraph text) | 100 | 216.6 | 217.7 | 1.005 |
| single set | 1,000 | 422.3 | 419.8 | 0.994 |
| single set | 10,000 | 2414.9 | 2404.0 | 0.995 |
| batch bold ×1 | 1,000 | 426.8 | 425.5 | 0.997 |
| batch bold ×10 | 1,000 | 434.0 | 440.5 | 1.015 |
| batch bold ×100 | 1,000 | 535.7 | 595.8 | 1.112 |
| batch bold ×1,000 | 1,000 | 1509.8 | 2101.2 | 1.392 |
| batch bold ×1,000 | 10,000 | 3527.0 | 4104.3 | 1.164 |
| batch bold ×10,000 | 10,000 | 24432.3 | 30435.3 | 1.246 |

Against the PERFORMANCE BAR:

- **Open**: within +1 µs (≤ 5 % or ≤ 20 µs) — met.
- **Single-op write**: within ±1 % (≤ 10 %) — met.
- **Batch**: 1k-op homogeneous batch 1.39× (1,000-paragraph document) and
  1.16× (10,000) of one hand-written leaf `Edit` doing the same work (≤ 2×) —
  met; 10k ops 1.25×. The batch first measured 3.76×: each set resolved
  `/body/p[N]` by building a node per skipped sibling; the ordinal index of
  §4 fixed it.
- **Query**: no numeric budget; the unified walk costs about 0.18 µs per
  matched paragraph against 0.03 µs for a raw leaf index walk (≈ 6×), the
  price of node construction, compiled-plan dispatch and per-match paths.
  Recorded as a follow-up, not a regression (ODT had no query before).
- **Scaling note.** The leaf 10k batch (24.4 ms) grows faster than ten times
  the 1k batch on the same document: `XmlSplicePublication::replace` checks
  each new splice against all earlier ones (O(n²), `litchi-odf-common`, not
  owned here).

## 7. Gaps and follow-ups

Listed in `capabilities()` (contextual operation gaps) unless noted:

- **Table grid.** Cells cannot be added, removed or moved; rows with vertical
  merges or covered cells cannot be removed or moved (`M6 ODT table grid
  edits`).
- **Drawing and media.** Adding pictures and frames needs package media
  resources (`M6 ODT drawing and media edits`); frames and images read, frame
  geometry sets, both remove.
- **Common styles** in `styles.xml` are read-only; formatting writes derive
  automatic styles (`M6 ODT style definitions`).
- **Clearing formatting** with `null` is refused; set the off value (`M6 ODT
  formatting removal`).
- **Typed cell values.** Cell text writes string cells (`M6 ODT typed cell
  values`).
- **Signed/encrypted packages** stay exact-no-op only (0268 G5).
- **Flat ODF.** `.fodt` input is detected and refused with
  `unsupported_capability` (unchanged shared open path); the leaf flat
  snapshot has no body index or package owner to publish through
  (`M6 flat ODF owner`).
- **Element gaps** (read and write): header, footer, numbering, field, shape,
  chart, OLE, TOC, revision, column, equation, break, tab, group, formfield,
  settings — the body index does not model them yet (`M6 ODT body index`).
- **Effective formatting.** Properties are local (the node's automatic
  style); inherited paragraph-style formatting is not reported on runs
  (`r[bold=true]` matches no run of a bold paragraph style). Not a
  capability gap; same contract as DOCX.
- **Query speed** (§6) and the O(n²) splice-publication check — follow-ups.
- **0268 G2–G4** did not block this adapter and are unchanged:
  `MutableDocument` remains a public writer, scoped families keep their
  vocabulary, `SourceBackedDocument` keeps its name.
- **Shared-helper note.** No `format/odf_odt.rs` was needed: every ODF helper
  this adapter uses lives in `litchi-odt` or `litchi-odf-common`. The
  sibling ODS/ODP adapters may want the same per-handle ordinal index and
  splice-based writes; consolidate after the three land.

## 8. Gates (all exit 0, on the final commit)

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-odt -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo test --no-fail-fast -p litchi-odt --all-features` | 0 (1,111 passed) |
| `cargo test --no-fail-fast -p litchi-ops --all-features` | 0 (1,984 passed; capability/registry goldens and query-corpus differential included) |
| `cargo test --no-fail-fast -p litchi --all-features` | 0 (260 passed) |
| `cargo test --no-fail-fast -p litchi-cli` | 0 (28 passed) |
| `cargo test -p litchi-ops --no-default-features --features odt --test capabilities --test registry --test odt_unified_write` | 0 (single-feature goldens) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-odt -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 (26 profiles, 1,201 bans, 39 families) |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (254 tests) |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI> --check` | 0 |
