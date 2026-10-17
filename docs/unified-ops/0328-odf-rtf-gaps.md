# 0328 — ODT, ODP and RTF element gaps against the DOCX and PPTX vocabularies

- Status: implemented on `uo2/odf-rtf-gaps`, cut from `feat/unified-ops-api`
  at `fc274041da`. Trunk was merged twice: at `84732bed28` (records 0325
  and 0326) and at the merge that follows this record (record 0330).
- Date: 2026-10-17
- Input: the ODT, ODP and RTF rows of `officecli-capability-delta.md`.
  - ODT and RTF listed numbering, headers, footers, fields, the TOC,
    tabs, shapes, sections and bookmarks as gaps or missing against the
    DOCX vocabulary.
  - ODP listed hyperlinks, line breaks, columns, media, equations and
    chart parts against the PPTX vocabulary.
  - Records 0319 (DOCX P0/P1), 0321 (PPTX elements) and 0324 (DOC) are
    the starting point. This record also carries the section-break
    follow-up of record 0308.
- Owned:
  - `litchi-odt`, `litchi-odp` and `litchi-rtf`;
  - `litchi-ops/src/format/{odt,odp,rtf}*`;
  - the ODT/OTH/ODM, ODP and RTF rows of the capability and registry
    goldens, the CRUD checklist and the fidelity report.
  - No hunk in the shared dispatch files (`read`, `write`, `transaction`,
    `save`, `create`, `capabilities`, `file`) or in `litchi-ops/Cargo.toml`.
- `performance_claim`: none. §6 reports every new and touched path.

## 1. Result

Each format answers the DOCX or PPTX paths and property names of
records 0319, 0321 and 0324.

The parts are reached by path and query. Node views keep listing the
structural children they listed before, as DOC does (record 0324). The one
exception is a part's own parts: a numbering lists its definitions, and a
definition lists its levels.

### ODT, shared by OTH and ODM

| element | get / query | set | add | remove |
|---|---|---|---|---|
| `/header[N]`, `/footer[N]` (master-page regions, `@type=default/even/first`) | ✅ text, `type`, `masterPage` | ✅ `text`, published into `styles.xml`; a region holding a field or other content a text set cannot keep is refused | — | — |
| `/numbering`, `abstractNum[N]`, `num[N]`, `level[L]` (list styles) | ✅ `abstractNumCount`, `numCount`; definition `id`, `type`, `name`, `styleName`; instance `id`, `abstractNumId`; level `ilvl`, `start`, `format`, `lvlText`, `justification`, `suff`, `indent`, `hanging`, `font`, `size`, `color`, `bold`, `italic` | — | — | — |
| `/field[N]`, `/toc[N]` | ✅ `fieldType`, `name`, `format`, `fldLock`, `evaluated`, `element`, text; TOC `levels`, `title`, `hyperlinks`, `pageNumbers`, `name` | — | — | — |
| `/body/p[N]/tab[M]` | ✅ `pos` (twips), `val`, `leader` | ✅ | ✅ (position order) | ✅ |
| `/body/shape[N]` | ✅ `geometry`, `name`, `width`, `height`, `anchor.x`, `anchor.y`, `anchor`, `wrap`, `fill`, `line.color`, `line.width`, text | — | — | — |
| `/body/equation[N]` | ✅ `formula` (MathML annotation or text), `mode` (`display` or `inline`), `name`, `width`, `height` | — | — | — |
| `/body/tbl[N]/col[C]` | ✅ `width` | — | — | — |

Two leaf changes go with this:

- The leaf gains an on-demand part index (`litchi_odt::body::Parts`). It
  takes one pass over each of `content.xml` and `styles.xml`, and it never
  changes the body index.
- The new `Edit::set_region_text` publishes `styles.xml` splices.

Fixing `clone_style` also fixed an existing bug: a style that already had
`style:paragraph-properties` no longer gains a second set.

### ODP

| element | get / query | set | add | remove |
|---|---|---|---|---|
| `/slide[N]/shape[M]/hyperlink` (click listener or `draw:a`) | ✅ `link` (URL, `slide[N]`, the PresentationML jump names) | ✅ `link` | ✅ | ✅ |
| `…/paragraph[K]/run[L]/hyperlink` (`text:a`) | ✅ `link`, `tooltip` | ✅ `link`, `tooltip` | ✅ (wraps the run) | ✅ (keeps the text) |
| `…/paragraph[K]/linebreak[B]` | ✅ | — | ✅ | ✅ |
| `/slide[N]/table[M]/col[C]` | ✅ `width` | — | — | — |
| `/slide[N]/video[M]`, `/slide[N]/audio[M]` (`draw:plugin`) | ✅ `name`, `id`, `alt`, `source`, `x`, `y`, `width`, `height`, `zorder`, `mediaType`, `mimeType`, `loop`, `mute` | — | — | — |
| `/slide[N]/equation[K]` | ✅ `id`, `formula`, `name`, `x`, `y`, `width`, `height` | — | — | — |
| `/slide[N]/chart[M]/series[K]`, `…/axis[@role=ROLE]` | ✅ series `values`, `nameRef`, `chartType`, `axis`, `pointCount`; axis `role`, `title`, `dimension`, `name`, `categories`, `style` | — | — | — |

Links are refused for script, data, file and archive schemes (record
0321). Every change is a checked scene splice
(`litchi_odp::scene::Change::{SetRunLink, LinkRun, SetShapeClick,
InsertLineBreak, RemoveLineBreak}`).

### RTF

| element | get / query | set | add | remove |
|---|---|---|---|---|
| `/section[N]` | ✅ the DOC section keys (`type`, page size, orientation, the seven margins, `columns`, `columnSpace`, `columns.separator`, `titlePage`, `pageNumFmt`, `pageStart`, `direction`, `vAlign`) | ✅ `type`, page size, orientation, margins, `columns`, `columnSpace`, `titlePage`, `vAlign` | — | — |
| `/bookmark[N]`, `/bookmark[@name=X]` | ✅ `name`, text | ✅ `name` (both groups; duplicate names refused) | — | ✅ (its text stays) |
| `/field[N]`, `/toc[N]` | ✅ `fieldType`, `instruction`, `fldLock`, `evaluated`; TOC `levels`, `hyperlinks`, `pageNumbers`, `instruction` | ✅ `instruction`; TOC `levels`/`hyperlinks`/`pageNumbers` rewrite the `\o`, `\h` and `\n` switches | — | — |
| `/body/p[N]/tab[M]` | ✅ `pos`, `val`, `leader` | ✅ | ✅ (position order) | ✅ |
| `/body/textbox[N]` (text-box shapes, legacy `\dptxbx`) | ✅ text, extent | — | — | — |
| `/body/shape[N]` (`\shp`) | ✅ extent, anchor offset, `wrap`, `behindDoc`, `geometry`, `name` | — | — | — |

The writes are new leaf mark edits:

- `litchi_rtf::edit::Edit::{set_field_instruction, rename_bookmark,
  remove_bookmark, insert_bookmark, set_paragraph_tab_stops,
  set_section_values}`, with `SectionValue`.
- Each edit tokenizes the exact source once and splices only the bytes
  that spell the entry:
  - the `\fldinst` text;
  - the `\bkmkstart` and `\bkmkend` groups;
  - the paragraph's `\tx`/`\tb` controls with their alignment and leader
    controls;
  - the section's `\pgwsxn`-family controls.
- The candidate is reparsed with the source's open policy. It must read
  back as staged, with every other paragraph, run, field, bookmark, section,
  table and drawing unchanged.
- The result publishes as a `mark.transfer` root transfer, which replays
  durably.

**Section breaks (follow-up of record 0308).** Producer paragraph splicing
used to refuse any body that had a section break. A root-level `\sect`,
with the section controls after it up to the next paragraph reset, text
or group, is now a block of its own that keeps its place among the
paragraphs, as a table row does. Every section must read back unchanged:

- a paragraph moves across a section break and back byte-exactly;
- removing a section's only paragraph is refused, because the section
  would read differently.

### Tests and fixtures

- **Tests:**
  - `crates/litchi-ops/tests/{odt,odp,rtf}_unified_parts.rs` (8, 7 and 6
    tests);
  - `crates/litchi-rtf/tests/marks.rs` (7);
  - the leaf tests appended to `litchi-odt/src/body/tests.rs` and
    `litchi-odp/src/scene/tests.rs`.
- **Changed existing assertions:**
  - `rtf_unified_read` now expects `/body/p[1]/shape[1]` to answer
    `unsupported_path`, because shapes are body children; the equation
    gap stays.
  - `odp_unified_read` now expects `path_not_found` for a missing video
    and `unsupported_capability` for animation.
- **Fixtures** in `test-data/unified-ops/0294/` (each with a row in its
  `PROVENANCE.md`):
  - `parts.odt`, `parts.oth` and `parts.odm` (3,212, 3,219 and 3,223
    bytes);
  - `parts.odp` (3,196 bytes);
  - `parts.rtf` (1,028 bytes, hand-written).
  - They are written by `tests/acceptance_fixtures.rs`.

## 2. Gaps and deferrals

Unchanged gaps are still declared with a follow-up in the capability
report:

| subject | formats | follow-up |
|---|---|---|
| `revision` (P2) | ODT, RTF | revision nodes and accept/reject |
| `formfield` writes | ODT, RTF | form-field nodes (FORMTEXT, FORMCHECKBOX, FORMDROPDOWN) and their writes |
| `sdt` | ODT, RTF | DOCX-only (`unsupported_path`) |
| `watermark` | ODT, RTF | the header WordArt shape |
| `equation` | RTF | EQ fields and math zones as equation nodes |
| `ole` | RTF | embedded objects as `ole` nodes |
| `chart`, `ole`, `break`, `group`, `settings` | ODT | as declared in the ODT spec |
| `animation` | ODP | PPTX animation timing |
| `theme`, `modernComment`, `model3d`, `zoom` | ODP | PPTX-only (`unsupported_path`) |

New follow-ups from this record:

1. **Header and footer formatting reads and region add/remove (ODT).**
   Today only the text and identity keys are read.
2. **Writes to ODT list levels and table columns, ODP table columns, and
   ODT shapes and equations.** All of these are read-only today.
3. **ODP shape click tooltips.** A click listener has no title, so a
   shape link's `tooltip` set is refused; a run link's tooltip is
   writable.
4. **RTF bookmark add.** The leaf `Edit::insert_bookmark` exists and is
   tested. The unified `add` needs bookmarks parented under paragraphs,
   whereas the spec parents them under the document as DOCX does.
5. **RTF field add/remove and TOC add.** Instructions are writable; field
   groups are not created or removed.
6. **Flat `.fodt` second edits.** A second formatting edit on a flat ODT is
   still refused. This is unchanged and reported as one-way/refused by
   fidelity.
7. **Read misses** (§6): point reads of parts pay the generic path, node
   and projection cost.
8. **Child-listing overhead** (§6): node views of the root, body,
   paragraphs and slides are 3–19% slower.

## 3. Coverage

**CRUD checklist** (`docs/unified-ops/results/0289/crud-checklist.md`):

| format | elements | claimed cells | ok | FAIL |
|---|---:|---:|---:|---:|
| odt (oth, odm alike) | 22 → 34 | 205 → 288 | 205 → 288 | 0 |
| odp | 30 → 30 | 170 → 228 | 170 → 228 | 0 |
| rtf | 21 → 28 | 154 → 214 | 154 → 214 | 0 |

**Fidelity** (`fidelity.md`) adds three structural round trips: `tab
add+remove` for ODT, OTH and ODM, and `linebreak add+remove` for ODP. RTF
keeps its `paragraph move+back`. `parts.rtf` reads back byte-exact, and
moves across its section break.

**OfficeCLI delta** (`officecli-capability-delta.md`):

| format | elements supported | gap | missing | reads supported | writes supported |
|---|---|---|---|---|---|
| odt | 19 → 31 | 12 → 5 | 13 → 8 | 63 → 112 | 12 → 15 |
| odp | 21 → 28 | 8 → 5 | 4 → 0 | 97 → 114 | 9 → 11 |
| rtf | 22 → 29 | 8 → 5 | 14 → 10 | 68 → 109 | 12 → 15 |

**Goldens.** The capability and registry goldens were re-blessed in both
feature sets: all features, and every feature except `encryption`. Both
sets produce the same files.

**Query corpus.** The ODT and ODP corpus digests changed only because new
fixture files were added. Moving the fixtures aside gave the old
digests. RTF is not in the corpus. The OfficeCLI script corpus covers DOCX
only.

## 4. Glossary

`docs/API_GLOSSARY.md`, section "ODT, ODP and RTF element parts (record
0328)", lists every new public leaf API and every new unified kind. The
glossary gate passes.

## 5. Byte identity

- Existing outputs are byte-identical. The CRUD rows for existing elements
  and the fidelity rows for existing fixtures did not change, apart from
  the new rows above.
- The new writes keep unedited bytes exact; the tests assert one-splice
  diffs:
  - ODT region and tab splices;
  - ODP link and line-break splices;
  - RTF `{ NUMPAGES }` → `{ SECTIONPAGES }`, the TOC switches, both
    bookmark groups, tab and section controls.
- Setting what an element already holds publishes the source bytes
  unchanged (tab stops, ODP links, ODT header text).

## 6. Measurements

**Setup.** Release builds, `taskset -c 27`, on a shared 32-core host (load
about 4). Evidence is in [results/0328/](results/0328/). Binary SHA-256s
and build commands are in `binaries.sha256`; `cargo fmt` later reflowed
the bench source without changing code.

**New paths, unified against leaf.** The harness is
`examples/odf_rtf_parts_bench.rs`, run on the `parts.*` fixtures.

- Each write times one edit against the same opened owner. The leaf leg
  runs the leaf verb and commits. The unified leg runs the `File` verb,
  with path resolution, admission and the returned path.
- Both legs publish identical bytes; this is checked outside the clock.
- Reads are warm.
- Each trial is 9 ABBA rounds after 2 warm-ups, with 64 iterations a
  sample; the figures are medians.
- There are three trials (`parts-trial-{1,2,3}.json`).

| case | leaf µs | unified µs | unified/leaf |
|---|---:|---:|---:|
| `odt_footer_text_set` | 248.8–250.8 | 250.3–252.0 | 1.005–1.006 |
| `odt_tab_add` | 251.4–252.4 | 255.0–255.5 | 1.010–1.016 |
| `odp_link_set` | 519.7–522.6 | 522.0–526.5 | 1.004–1.007 |
| `odp_linebreak_add` | 520.8–524.3 | 524.0–526.3 | 1.004–1.006 |
| `rtf_field_instruction_set` | 30.4–30.8 | 30.9–31.2 | 1.007–1.017 |
| `rtf_bookmark_rename` | 30.4–30.7 | 30.9–31.1 | 1.008–1.018 |
| `rtf_tab_add` | 38.9–39.1 | 39.2–39.7 | 1.009–1.016 |
| `rtf_section_set` | 31.2–31.2 | 31.6–31.9 | 1.011–1.019 |
| `odt_header_get_format` (read) | 0.09 | 0.29–0.30 | **3.22–3.26** |
| `odp_link_get_format` (read) | 0.10 | 0.66–0.68 | **6.47–6.86** |
| `rtf_section_get_format` (read) | 0.00 | 0.22–0.24 | **354–504** |
| `rtf_tab_get_format` (read) | 0.01 | 0.24–0.27 | **46–54** |

- **Writes.** Every write is within 2% of the leaf, well inside the 10%
  single-op budget.
- **RTF mark edit cost.** An RTF mark edit costs 30–39 µs, against
  250–520 µs for the ODF writes. Most of it is the reparse that proves
  the candidate.
- **Four read misses, not fixed.** The unified point reads cost a fixed
  0.2–0.7 µs:
  - the path is parsed and resolved;
  - a node is built;
  - every key of the element's table is projected (a section has 19 keys,
    a link 2 after its click scan).

  The leaf legs instead read one typed value; `rtf_section_get_format`'s
  leaf leg borrows one struct in about 1 ns. This is the cost shape of
  record 0324's `section_get_format`. Follow-up 7.

**Existing paths, after against before.** The harness is
`examples/odf_rtf_view_regression.rs`. It uses the public `File` API only,
on `writer-paragraph-styles.odt`, `impress-basic.odp` and
`rtf/testDefaultEncodingParse.rtf`.

- It was built from a `git archive` of `fc274041da` and from the branch.
- Each run is 15 rounds after 3 warm-ups. There were two runs, B A A B B A
  and A B B A A B (A is before, B after); see `view-regression.txt`.
- Instruction counts per iteration from callgrind are in `callgrind.txt`.

| scenario | before ns | after ns | after/before (run 1 / run 2) | instructions after/before |
|---|---:|---:|---:|---:|
| ODT `/body/p[1]` format | 1,885–1,918 | 1,704–1,731 | 0.908 / 0.904 | 0.931 |
| ODT query `p` + every format | 34,640–34,893 | 30,338–30,762 | 0.880 / 0.881 | 0.854 |
| ODT root children | 164–195 | 172–222 | 1.052 / 1.099 | 1.063 |
| ODT body children | 360–465 | 396–532 | 1.126 / 1.191 | 1.112 |
| ODT `/body/p[1]` children | 210–223 | 217–249 | 1.009 / 1.112 | — |
| ODT text view | 1,039–1,081 | 1,073–1,114 | 1.021 / 1.043 | — |
| ODT `/body/p[1]` text set | 672,545–682,309 | 689,960–694,050 | 1.020 / 1.021 | — |
| ODP slide children | 304–344 | 328–352 | 1.070 / 1.124 | 1.014 |
| ODP shape children | 300–330 | 308–333 | 0.990 / 1.072 | — |
| ODP paragraph children | 385–397 | 394–421 | 1.008 / 1.061 | — |
| ODP shape format | 580–596 | 600–631 | 1.034 / 1.036 | 1.015 |
| ODP query `shape` + every format | 3,634–3,706 | 3,697–3,748 | 1.000 / 1.015 | — |
| ODP text view | 204–219 | 201–210 | 0.995 / 0.949 | — |
| RTF `/body/p[1]` format | 224–239 | 220–238 | 0.974 / 0.987 | — |
| RTF root children | 126–150 | 135–156 | 1.046 / 1.109 | 1.067 |
| RTF body children | 157–177 | 165–188 | 1.075 / 1.108 | 1.044 |
| RTF `/body/p[1]` children | 202–228 | 212–237 | 1.034 / 1.063 | — |
| RTF query `p` + every format | 410–438 | 448–468 | 1.030 / 1.056 | 1.069 |
| RTF text view | 32–33 | 33 | 1.000 / 1.000 | — |
| RTF `/body/p[1]` text set | 2,798–2,956 | 2,814–2,856 | 0.977 / 1.018 | — |

The first measurement found larger regressions. Callgrind traced them, and
four fixes landed in `df1428b75f`:

- **ODT attribute scan.** The ODT attribute scanner's `str::find` closure
  had fallen out of line once the part index added callers, and every ODT
  attribute read paid a call. It is now an inline byte scan. ODT format
  reads are 9% faster and paragraph queries 12% faster than the base.
- **ODT child-listing struct.** The ODT child listing counts only the
  structural kinds and boxes its part tail.
- **RTF tab stops.** RTF paragraph listings read tab stops from the
  cursor's paragraph. Before, each listing walked the body to find the
  paragraph, which was quadratic over a query.
- **ODP property values.** ODP property values borrow the object reference
  and box link values. Object property reads are inlined again.

**Child-listing miss, not fixed (follow-up 8).** Node views of the root,
body, paragraphs and slides are 3–19% slower, which is 10–60 ns.

- Each listing now checks for part kinds and iterates longer child
  tables. In instructions this is +1–11%.
- Callgrind attributes the rest to one more struct move per listed child.
  Inlining the constructors made it worse, so they were left out of line.
- Queries and writes over the same documents are unchanged or faster.

No claim of faster ODT, ODP or RTF reads or writes is made beyond the ODT
attribute scan above.

## 7. Gates

Each command was run from the worktree with the build environment sourced.
Every exit status was checked, not a piped tail.

- `cargo fmt --all --check` → 0 (also re-run after each trunk merge).
- After the second trunk merge (record 0330, `d761892de3`), these were
  re-run and each → 0, with no file changing:
  - the registry, capabilities and `unsupported_warn` tests;
  - both golden blesses, the CRUD and fidelity blesses and the delta;
  - `cargo clippy -p litchi-ops -p litchi --all-targets --all-features`;
  - every litchi-ops test (`--lib`, then the 150 targets in groups).
- `cargo clippy -p litchi-odt -p litchi-odp -p litchi-rtf -p litchi-ops -p
  litchi -p litchi-py --all-targets --all-features -- -D warnings` → 0.
- `cargo clippy -p litchi-ops --no-default-features [--features
  odt|odp|rtf] --all-targets -- -D warnings` → 0 for all four.
- `cargo test --all-features` for each of `-p litchi-odt`, `-p
  litchi-odp` and `-p litchi-rtf` → 0.
- `cargo test -p litchi-ops --all-features`: `--lib` (929) → 0, `--doc` →
  0. The 150 integration targets ran in ten groups of 15, each → 0, and
  each group's executables were deleted afterwards.
- `cargo test -p litchi --all-features` → 0.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p
  litchi-odt -p litchi-odp -p litchi-rtf -p litchi-ops` → 0.
- `python3 tools/check_crate_boundaries.py` → 0; `python3
  tools/non_iwork_gate.py verify` → 0.
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir
  <worktree>/target/glossary` → 0.
- `crates/pyo3-litchi/tests/run_tests.sh` → 0 (283 tests).
  - The first run found that the CRUD mirror adds a tab stop with a
    placement, which RTF refused.
  - RTF now keeps position order, as ODT does (`313a2b3707`).
- `cd tools/perf-baseline && cargo check --locked --all-targets` → 0.
- Goldens: `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test
  capabilities --test registry` → 0. The same command with `--features`
  set to every litchi-ops feature except `encryption` → 0.
- `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test
  crud_matrix --test corpus_fidelity` → 0.
- `python3 tools/officecli_capability_delta.py --officecli
  <OfficeCLI checkout>` → 0.
- `python3 tools/test_no_3rdparty_references.py` → 0; `python3 -m
  unittest tools.test_officecli_capability_delta` → 0.
