# 0290 — Phase G checklist failures and metadata `language`

- Status: implemented on `uo2/checklist-fixes`, cut from
  `feat/unified-ops-api` at `99cc3138eb`. Follows record
  [0289](0289-phase-g-acceptance-evidence.md) §1.1 and §6 item 2, and the
  read gap of record [0288](0288-odf-follow-ups.md) §1.
- Date: 2026-10-15
- Owned: small fixes in separate per-crate commits (`litchi-ops` query
  walker, `litchi-drawingml`, `litchi-ppt`, `litchi-odf-common` +
  `litchi-odt`, `litchi-rtf`), the `crud_matrix` and `corpus_fidelity`
  tests and their goldens, the new `tests/checklist_fixes.rs` and
  `examples/checklist_fixes_bench.rs`.
- `performance_claim`: none for the API. The touched paths are measured
  before/after (§5): no change beyond noise.

## 1. Result

Record 0289 left nine claimed-supported cells failing (F1–F7; F5 was
closed by record 0288's ODC writes before this base). This record fixes
six of the seven remaining entries of `KNOWN_FAILURES`; F3 stays refused
for the reason in §3.

| # | cell | fix | crate |
|---|---|---|---|
| F1 | pptx and odp `placeholder` query | bare `placeholder` compiles to `shape` plus a placeholder test | `litchi-ops` (query) |
| F2 | pptx `chart` view | negative `c:axId`/`c:crossAx` read as the same 32 bits | `litchi-drawingml` |
| F3 | pptx `ole` remove | **kept refused** (§3) | — |
| F4 | ppt `slide` set `hidden` | missing `SlideShowSlideInfoAtom` is inserted | `litchi-ppt` |
| F6 | oth `metadata` set | ODT body edit creates `meta.xml` on demand | `litchi-odf-common`, `litchi-odt` |
| F7 | rtf `picture` remove | removal no longer inspects the payload's format | `litchi-rtf` |
| — | `File::metadata()` `language` (0288 §1) | shared `meta.xml` conversion maps `dc:language` | `litchi-odf-common` |

Checklist after the fixes (only changed rows; claimed = ok + FAIL +
unreached + cells not exercised):

| format | ok | FAIL | unreached | move/swap ok |
|---|---|---|---|---|
| pptx | 178 → 180 | 3 → 1 | 8 | 2 |
| ppt | 65 → 68 | 1 → 0 | 8 | 2 |
| rtf | 153 → 154 | 1 → 0 | 0 | 2 |
| odp | 165 → 166 | 1 → 0 | 4 | 4 → 6 |
| oth | 177 → 180 | 1 → 0 | 21 | 26 |

ppt `slide` and oth `metadata` also gain their `edit` and `batch` cells,
which replay the now-successful `set`; odp `placeholder` gains `move` and
`swap` (the walker now finds a pair of placeholder siblings).

## 2. Fixes

### F1 — bare `placeholder` selects placeholder shapes (`litchi-ops`)

**Vocabulary decision.** The registry lists `placeholder` as an element of
PPTX, PPT and ODP, and paths address placeholders with `placeholder[N]`,
`placeholder[TYPE]` and `@type/@idx`. Selectors keep that element: a bare
`placeholder` selects every placeholder. What differs is how the format
walks them:

- **PPT** walks a placeholder as its own semantic kind already
  (`Inspect::kind` returns `Placeholder`), so the selector worked.
- **PPTX and ODP** walk a placeholder as a `shape` (OfficeCLI parity:
  `shape` matches title and body placeholders). The type selector compiled
  to the `placeholder` kind, which no walked node carries, so it matched
  nothing.

`Inspect` gains `PLACEHOLDER_SHAPES` (default `false`) and
`is_placeholder(node)`. When a format sets the constant, the compiler maps
a bare `placeholder` to the `shape` kind plus a new structural test,
`Structural::Placeholder`. The results are the same shapes `shape`
matches, at their `/shape[N]` addresses, filtered to:

- PPTX: a `p:sp` with a `p:ph` binding (`Node::is_placeholder`), title
  placeholders included;
- ODP: a shape with a `presentation:class` (the existing semantic kind).

`shape` itself is unchanged. `shape[type=placeholder]` misses title
placeholders (their type is `title`); the bare selector does not.

Cost: one constant check while compiling a `placeholder` element; the
walker's structural-test match gains one arm that only `placeholder`
plans reach.

Goldens: `tests/golden/query_corpus/digests.json` changes in one key,
`pptx placeholder`, from 0 to 4,644 matches (masters and layouts hold most
of them). Re-blessed with `LITCHI_BLESS_GOLDEN=1` after checking that no
other key moved.

### F2 — negative chart axis IDs (`litchi-drawingml`)

The schema types `c:axId` and `c:crossAx` as `ST_UnsignedInt`, but
PowerPoint writes IDs at or above 2^31 as negative signed integers. Both
`ooxml/pptx/radar-chart.pptx` and `ooxml/pptx/line-chart.pptx` do, e.g.
`-1884094432`.

**Recorded leniency:** `chart::axis::parse_axis_id` accepts a `u32`, or an
`i32` read as the same 32 bits (`i32::cast_unsigned`). It serves:

- every axis-ID read: the ten type-group `c:axId` sites, the axis's own
  `c:axId` and its `c:crossAx`;
- the source-preserving editor's axis lookup.

Out-of-range values and non-numbers are still `invalid_attribute`.

**Writes stay schema-valid.** The typed writers spell IDs as unsigned
decimals, so a rewritten chart carries `2410872864` for `-1884094432`.
Exact-splice edits keep the producer's bytes.

Tests:
- `chart::reader::tests::reads_negative_producer_axis_ids_as_the_same_bits`:
  reads the IDs, writes them unsigned, re-reads, and refuses
  `-2147483649`, `4294967296` and `x`;
- `chart::edit::tests::edits_find_axes_with_negative_producer_ids`.

### F4 — inserting the slide-show information atom (`litchi-ppt`)

`set_slide_hidden` rewrote only an existing `SlideShowSlideInfoAtom`.
Producer slides and slides from the PPT writer (without a transition or
timing) have none. The change:

- **Read.** `slide_hidden` reads a slide without the atom as shown
  (`false`). [MS-PPT] 2.5.1 makes the atom optional and `fHidden` clear by
  default. Several atoms, or a malformed one, are still a refusal.
- **Hide.** Hiding such a slide inserts one atom directly after the
  container's `SlideAtom` ([MS-PPT] 2.5.1 order). The atom carries the PPT
  writer's default slide timing: no transition, no sound, manual advance,
  medium speed. Then `fHidden` is set. A container whose first child is
  not a `SlideAtom` is refused.
- **Checks.** The rewrite check expects the record to grow by exactly 24
  bytes, and the existing round-trip check (document model unchanged,
  `slide_hidden` reads back) still runs.
- **Show.** Showing a slide without the atom is an exact no-op. Showing it
  later clears the flag and keeps the atom. The exact inverse of the
  hiding commit restores the source bytes.
- **Durable replay.** It goes through the same function, so a durable
  patch inserts the atom too.

Visible side effect: once inserted, the slide reports a transition-info
record (`Slide::transition()` returns `Some` with type `None`, on-click
advance) where it returned `None`. The unified API does not project
`transition()`; the slide view gains a `hidden` value it lacked.

Test: `slide_order::tests::slide_visibility_inserts_a_missing_slide_show_information_atom`.
It covers the atom position and count, the untouched second slide and
text, the advance read, the exact patch and its inverse, and show after
hide.

### F6 — `meta.xml` on demand for OTH and ODM (`litchi-odf-common`, `litchi-odt`)

OTH and ODM files open through the ODT adapter. Their `/metadata` set
runs the ODT body edit, which planned `meta.xml` splices only and refused
a package without the part. The 0288 engine already had the creation
route (`Assignments::created_part`, used by its generic `Edit`), so this
record wires the ODT body publication through it:

- **Engine.** `created_part` becomes the public
  `Assignments::created_package`. The new `Assignments::plan_publication`
  reports `Publication::MissingPart` with a single `meta.xml` lookup
  (§5).
- **ODT body publication.** When metadata is staged and the package has no
  `meta.xml`, it publishes the content splices first (if any), then adds
  the created part. The part declares the content's `office:version`, and
  every other member is copied.
- **Removals only.** Staged removals alone write nothing.

Edits that stage no metadata look nothing up.

Test: `crates/litchi-odt/tests/metadata_without_meta_xml.rs` covers
creation (`office:version` and language), a second edit on the created
part, the untouched `content.xml`, and the removal-only no-op.

### F7 — picture removal ignores the payload format (`litchi-rtf`)

LibreOffice `n192129.rtf` stores GIF data under `\pngblip`. Before, the
standalone-picture locator checked every picture's payload signature,
even for whole-group removal, which moves the group's bytes unchanged.

The locator now takes a purpose, `Locate::{Payload, Group}`:

- **Payload.** Replacement keeps the PNG/JPEG signature check.
- **Group.** Removal, the durable removal patch, and the inverse insertion
  with its candidate check skip it.

Everything else still applies: the structural refusals (binary, nested,
fields, shapes, objects, compatibility wrappers, unknown syntax,
protection), the `MAX_EDITABLE_PICTURE_BYTES` bound, and
`verify_removal_candidate`.

Test: `picture_removal_crud::removes_groups_whose_payload_format_payload_edits_refuse`.
It removes a format-mismatched group and an EMF group exactly, restores
both through the durable inverse, and payload replacement still refuses.

### Metadata `language` (record 0288 §1 gap)

`litchi_odf_common::core::metadata::Metadata` read `dc:language`, but both
conversions to `litchi_core::Metadata` dropped it: `try_into_core` and the
`From` impl. Both now carry it.

This is the same conversion the 0288 agent reverted to keep the facade
Markdown golden fixed. That golden is our artifact, so it is updated here.

Test: `core::metadata::tests::common_metadata_carries_the_language`.

#### Facade Markdown golden (`crates/litchi/tests/golden/markdown_corpus.tsv`)

Six rows change, each in the `metadata+html` column (path and bytes
routes) only. The YAML front matter gains one `language:` line, and the
byte counts confirm it is the only change. The TSV header records the
change.

| fixture | `dc:language` | old → new | Δ bytes |
|---|---|---|---|
| `libreoffice-core/xmloff/qa/unit/data/tdf156707_text_form_control_borders.odt` | `en-CA` | `362:0b3b0f82ba922f83` → `378:db7aa3cf53e332bf` | +16 = `language: en-CA\n` |
| `odf/corpus/impress-master-layouts.odp` | `fr-FR` | `242:ba10f9cf7d88bf15` → `258:ff54e463c03c0f9e` | +16 = `language: fr-FR\n` |
| `odf/corpus/writer-user-fields.odt` | `de-DE` | `322:72651d3d0b03fd85` → `338:082d3757336d8abe` | +16 = `language: de-DE\n` |
| `odf/native-resave/libreoffice-resaved/tdf169979-litchi.odp` | empty `<dc:language/>` | `372:4d67cfd6e09affa2` → `385:428bb011dc0ca65e` | +13 = `language: ''\n` |
| `odf/native-resave/litchi-changed/tdf169979-litchi.odp` | empty | `321:cf8caeca88945894` → `334:48f00101a5d1d870` | +13 |
| `odf/odp/tdf169979.odp` | empty | `286:d821d451d756e1e1` → `299:78d4cbf0bcb5dac5` | +13 |

**ODT rows.** Two of the rows are ODT, which the 0288 note did not list.
The ODT facade Markdown reads metadata through the same conversion.

**Empty languages.** The three `tdf169979` packages (OpenXML/ODF
Translator output) carry an empty `<dc:language />`. They now report an
empty language, as they already report their empty `dc:creator` as an
empty `lastModifiedBy`. The conversion does not invent absence.

## 3. F3 kept refused: PPTX OLE removal under markup compatibility

**The fixture.** In every corpus OLE object (e.g.
`libreoffice-core/sd/qa/unit/data/pptx/ole.pptx`), the
`mc:AlternateContent` sits inside the `p:graphicFrame`, under
`a:graphicData`; it does not wrap the frame. Removing the whole
`p:graphicFrame` would remove the wrapper with it, and that range is
lossless.

**Why it stays refused.** The refusal comes before any shape logic. The
source-backed slide guard (`litchi-pptx`
`presentation::transition::validate_source` → `locate`) refuses any
markup-compatibility element outside a direct transition owner. Every
source-backed slide path runs that guard:

- the slide snapshot every edit starts from (`slide_snapshot_for`);
- the retained-catalog publication check
  (`validate_retained_raw_slide_snapshot`);
- the apply checks of the primary transaction (`durable::validate_pair`,
  `topology` after-images).

`stage_shape_removal` also computes shape offsets on the MCE-preprocessed
scene, which is rewritten for these slides, so its offsets do not map
onto the source.

**What a fix needs.** Removal needs a narrower admission that all four
layers agree on: MCE allowed only strictly inside top-level `p:spTree`
children, and a removal proven on the source bytes. The proof would
remove the k-th source child, preprocess both sides, and compare the
results. Other edits on such slides would stay refused. That is a
cross-layer change to the PPTX source-backed proof system, not a small
leaf fix.

**Follow-up:** "0290 PPTX removal under enclosed markup compatibility".
The `KNOWN_FAILURES` entry now cites this section.

## 4. Goldens and fidelity report

Each golden was re-blessed after its diff was reviewed:

- **`results/0289/crud-checklist.{md,tsv}`** (`LITCHI_BLESS=1 … --test
  crud_matrix`): only the rows in §1 changed. Changes the fixes cause
  elsewhere in the TSV:
  - pptx `series`/`axis` read cells now list `radar-chart.pptx` first,
    with 6 and 8 instances;
  - odp `placeholder` cells now start at `odf/odp/tdf169979.odp`, with 5
    instances.
- **`results/0289/fidelity.{md,tsv}`** (`--test corpus_fidelity`): three
  rows changed.
  - `ooxml/pptx/line-chart.pptx` and `radar-chart.pptx` (negative axis
    IDs): `none` → `refused`. The chart nodes now read, so the
    set-then-revert probe reaches `chartType`, a text property, and the
    generated text is refused as `invalid_value: unknown chart type`.
    That is a typed refusal of the probe value. Their no-op saves stay
    exact.
  - `tdf156707_text_form_control_borders.odt`: the probe now finds the
    `/metadata language` text before `/body/p[1]` text; still
    value-exact.

  PPTX revert counts move from refused 50 / none 26 to 52 / 24. Every
  no-op save is still byte-identical (771 exact, 0 differ).
- **`tests/golden/query_corpus/digests.json`**: §2 F1.
- **`markdown_corpus.tsv`**: §2.

## 5. Measurements

The example `checklist_fixes_bench.rs` was run on two builds of
`cargo build --release -p litchi-ops --example checklist_fixes_bench
--features pptx,odt,odp`:

- **before:** the base `99cc3138eb`, exported with `git archive` plus this
  example;
- **after:** this branch.

Both were pinned to one core of a shared host, ABBA × 2, with 21 timed
rounds after 3 warm-ups. The table gives medians of the per-leg medians,
in µs per call.

| scenario | run 1 before → after | run 2 before → after | after/before |
|---|---|---|---|
| `pptx_query_shape` (`shapes.pptx`) | 8.97 → 8.12 | 8.69 → 8.20 | 0.905, 0.944 |
| `pptx_chart_view` (`tdf114821.pptx`) | 304.6 → 298.3 | 301.9 → 302.0 | 0.979, 1.000 |
| `odt_metadata_set_publish` (`writer-definition-lists.odt`) | 745.8 → 723.5 | 754.3 → 722.1 | 0.970, 0.957 |
| `odp_metadata_set_publish` (`impress-master-layouts.odp`) | 1,011.8 → 962.5 | 1,017.2 → 965.5 | 0.951, 0.949 |
| `odp_metadata_read` (same ODP) | 0.0319 → 0.0387 | 0.0312 → 0.0385 | 1.213, 1.234 |

**A regression found and fixed on the way.** The first after-build checked
`package.has_file("meta.xml")` before planning the publication, which
checks it again. That lookup costs about 300 k instructions per call on
this package: `has_file` reparses the manifest. callgrind showed +1.9 %
instructions for `odt_metadata_set_publish` (1,882,144,389 →
1,918,121,826), and ABBA showed 1.03–1.04.

The fix adds `Assignments::plan_publication`, which looks the part up
once and returns `Publication::{Unchanged, Splices, MissingPart}`. The ODT
body publication uses it, and so does the shared engine's own commit,
which had the same double lookup since record 0288. ODS, ODP and the
object formats therefore get faster too: `odp_metadata_set_publish` is at
0.95.

After the fix, callgrind of `odt_metadata_set_publish` counts
1,882,235,832 instructions before and 1,881,685,047 after (−0.03 %).

**Other rows.**
- **`odp_metadata_read`** (+7 ns): `File::metadata()` returns a clone of
  the cached metadata, which now carries one more string (`fr-FR`). That
  is the requested output, not overhead.
- **Query and chart view:** within noise. The walker change adds no work
  outside `placeholder` plans. The axis-ID parse tries `u32` first, as
  before.

**Evidence.**
- [`results/0290/checklist-fixes-bench-abba.jsonl`](results/0290/checklist-fixes-bench-abba.jsonl).
- The binaries are not committed (AGENTS.md). Their SHA-256: before
  `c7d183fd601695d9a3585b4a5d46dc2714469ec15af5f30907d057a481b79e60`,
  after `22bacbc8397cd99359e84e100931a9db4f440e265a0cb911acc1363ba9e97a05`.
- Host: load 6–8, core 13.

## 6. Gates

All exit 0 on the final tree. The build environment had
`CARGO_BUILD_JOBS=6` and `RUST_TEST_THREADS=8`, and ran one cargo command
at a time.

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-{drawingml,ppt,rtf,odt,odf-common,ops} -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ods --lib --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features f] --all-targets -- -D warnings`, f ∈ none, pptx, odp, ppt, rtf, odt, oth, odm, ods | 0 each |
| `cargo test --no-fail-fast --all-features -p litchi-{drawingml,ppt,rtf,odf-common,odt,oth,odm,odp,odg,odc,odi,odb,odf-formula}` | 0 |
| `cargo test --all-features -p litchi-ods --test metadata_family` (the ODS suite needs a purged fixture) | 0 |
| `cargo test --no-fail-fast -p litchi-ops --all-features` (lib, doc, then the 121 integration tests in chunks of 25 to stay under the disk cap) | 0 |
| `cargo test --no-fail-fast -p litchi --all-features` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix` / `--test corpus_fidelity` | reviewed (§4) |
| `LITCHI_BLESS_GOLDEN=1 cargo test -p litchi-ops --all-features --test query_corpus_differential` | reviewed (one key, §2 F1) |
| `cargo test -p litchi-ops --test capabilities --test registry`, `--all-features` and every feature but `encryption` | 0, goldens unchanged (no bless needed) |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI>` | 0, table unchanged |
| `cargo test -p litchi-cli --all-features --test officecli_corpus`; `python3 tools/officecli_corpus.py --check` | 0, 0 (unchanged) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-{drawingml,ppt,rtf,odt,odf-common,ops}` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `python3 tools/test_no_3rdparty_references.py` | 0, 0, 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (267 tests) |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

The `litchi-ops` run found one test that pinned the old F4 gap:
`ppt_unified_write::absent_visibility_record_is_an_explicit_native_gap`.
It is replaced by `absent_visibility_record_is_inserted_when_hidden`,
which checks that hiding works, showing is an exact no-op, and the source
stays unchanged.

**New public items:**
- `litchi_drawingml::chart::axis::parse_axis_id`;
- `litchi_odf_common::metadata::{Assignments::created_package,
  Assignments::plan_publication, Publication}`;
- the crate-private `Inspect::{PLACEHOLDER_SHAPES, is_placeholder}`.

The glossary gate passes, and no glossary row changes.

## 7. Commits

- `f6ae012125`: F2 (`litchi-drawingml`).
- `19dea068a1`: F4 (`litchi-ppt`, plus the PPT unified test that pinned
  the gap).
- `da79f943e0`: F7 (`litchi-rtf`).
- `d81c4334e8`: F6 (`litchi-odf-common` metadata engine, `litchi-odt`).
- `dc1910cf9e`: metadata `language` (`litchi-odf-common` conversion,
  Markdown golden).
- `af2f2a1008`: F1 (`litchi-ops` query walker, query digest golden).
- The record commit: `KNOWN_FAILURES`, the re-blessed checklist and
  fidelity report, `tests/checklist_fixes.rs`,
  `examples/checklist_fixes_bench.rs` and this record.

## 8. Follow-ups

- F3: PPTX removal under enclosed markup compatibility (§3).
- 0289 §1.2's 70 unreached cells and §6's other items are unchanged.
