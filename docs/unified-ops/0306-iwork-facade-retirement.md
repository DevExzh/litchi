# 0306 — `litchi::iwork::Document` retired; iWork package directories and open options through `File`

- Status: implemented on `uo2/iwork-facade`, cut from `feat/unified-ops-api`
  at `e7fe30fb24`, trunk merged at `bc871c725c` (records 0303, 0304). Plan
  [0248](0248-coordinator-restart-plan.md) §6 Phase F, step 28.
- Date: 2026-10-16
- Owner decisions applied: **D1**, **D2** (2026-10-15).
- Commits: `4ae5372cef` (litchi-iwa-archive `bundle`, Pages on it),
  `40fa58ea0f` (litchi-keynote), `f38226b43e` (litchi-numbers),
  `36277adc96` (loose `Index/` bundles), `6b6efe04fd` (litchi-pages semantic
  prepared ingress), `4701b6a962` (litchi-ops), `d960232365` (facade
  deletion and every caller), `93f86b8b72` (single-pass directory open),
  `bc871c725c` (trunk merge), and the commit carrying this record and the
  bench matrix.
- In `crates/litchi-ops/src/format/iwork*` only directory and open-option
  plumbing changed (`iwork.rs`: `detect_bytes`, `capture_directory`,
  `open_directory`, `save_to`, `Owner::save_directory`;
  `iwork/error.rs`: two save variants and the limit classification), as
  agreed with the perf agent (record 0305), who edits the lazy trees.
- `performance_claim`: paired leaf/unified medians in §5 (release, one pinned
  core of a shared host, ABBA). No frozen-harness qualification.

Record 0296 kept the facade's read-only `litchi::iwork::Document` for four
gaps: Keynote and Numbers package directories, Pages/Keynote tables, text
roles and section kinds, and the missing iWork options in `OpenOptions`.
This record closes them, deletes the type and its module, and moves every
caller to `File`.

## 1. Package directories for every iWork format

| Layer | Change |
|---|---|
| `litchi-iwa-archive` | New public `bundle` module: `capture(root, CaptureLimits, Limits)` converts an app-authored bundle to its single-file form (the record 0299 algorithm, moved out of `litchi-pages`), `save(package, limits, path, …)` writes the bundle form through `publication::replace_directory_with`, plus `converted_limits`, `is_directory`, `CaptureError` (`From` into the archive `Error`), `CaptureLimitKind`, `INDEX_ZIP`. A bundle whose components sit loose in an `Index/` folder (the layout the facade reader accepted) converts too; one holding both layouts is refused. |
| `litchi-pages` | `package/directory.rs` delegates to `bundle`, keeping its limit and error mapping; behaviour unchanged. |
| `litchi-keynote`, `litchi-numbers` | `Package::open[_with_options]` accept a package directory (capture → converted limits → the ordinary exact-source ingress); `Package::save_directory`; `SaveError::{Bundling, Bundle}`. |
| `litchi-ops` | `open(dir)` captures the bundle once and opens the converted bytes through the in-memory detection, whose prepared container the leaf adopts (`Iwork::capture_directory`); caller leaf source limits, a legacy `index.xml`/`index.apxl`/`index.numbers` marker or a failed capture fall back to classification and the leaf's own capture, so errors stay the leaf's. The directory is the handle's origin; `save()` (and `save_as` onto an existing directory) writes the bundle back for all three formats (`Owner::save_directory`). |

Byte preservation: saving the directory fixtures back without an edit
reproduces **every bundle file byte for byte, `Index.zip` included**, for
Pages, Keynote and Numbers (`litchi-keynote` `native_fixture`,
`litchi-numbers` `document_reader`, `litchi-ops` `iwork_unified`
`directories::every_format_opens_a_package_directory_and_saves_it_back_in_place`).
After an edit, every untouched member (metadata, `Data/` media) keeps its
bytes; only `Index.zip` and the previews the leaf drops change. The corpus
fidelity report now lists the three bundles as corpus entries: open → empty
edit → `save()` in place on a copy, every file compared — 3 × `exact`
(`results/0289/fidelity.{md,tsv}`, re-blessed; corpus 977 → 980 entries).

## 2. iWork options in `OpenOptions`

`litchi_ops::OpenOptions::{pages_options, keynote_options, numbers_options}`
take the leaf `OpenOptions` (source and semantic limits), as `odf_options`
does for ODF.

- Bytes and file routes: detection prepares under the hard ceilings (as
  before); the detected format's semantic limits go to the prepared-source
  ingress (new hidden `litchi_pages::Package::__from_prepared_source_with_semantic_limits`;
  Keynote and Numbers already took them). Source limits other than the
  leaf defaults make the leaf parse the shared bytes again under them —
  a second container pass on that opt-in route only.
- Directory route: the leaf captures under the caller's options;
  `max_input_bytes` bounds the converted package.
- Errors: leaf archive limit errors, the Pages `Limit` and the Pages
  semantic text/section/storage bounds were classified `corrupt_file`; they
  are now `limit_exceeded` (`archive_code`, through the new
  `litchi-ops -> litchi-iwa-archive` edge, ADR 0024 amendment).
- Tests: `iwork_open_options_reach_the_leaves_on_every_route` (Pages
  semantic text limit, Keynote object limit, Numbers source limit, each on
  the bytes, file and directory routes; other formats unaffected; default
  options keep the exact bytes).

## 3. `litchi::iwork::Document` deleted

**Coverage check before deletion.** A scratch probe (not committed) opened
every iWork input under `test-data/` (125 files and the 3 bundles) with
both readers and looked for each facade text item (every `TextRole`: table
names and headers, cells, slide titles/content/additional/notes, section
headings/paragraphs/additional) in `File`'s text, annotated view or
outline: **128/128 without a gap**. Tables (Pages `tbl`, Keynote `table`),
headers, footers and footnotes are `File` elements since 0297–0301; section
kinds are the element kinds. Two facade reads have no equivalent, by
decision:

- the synthetic Numbers compatibility fixture's *detached* table model
  (`compatibility-oracles.hex`, model 5 attached to no sheet): not document
  content; `File` reads the rooted workbook. On that fixture `File`'s text
  is right, but its object tree refuses the source with `invalid_input`
  ("the selected Numbers table-merge source is invalid") — gap G2;
- the snapshot's own limit vocabulary (`SnapshotLimits`, `Resource`,
  `Stage`): replaced by the leaf options of §2.

**Deleted:** `crates/litchi/src/iwork/` (`Document`, `Snapshot`, `Summary`,
`Table`, `Slide`, `Section`, `SectionKind`, `TextRole`, `Text`, `Value`,
`CellView`, `Options`, `SourceLimits`, `SnapshotLimits`, `Error`,
`ErrorKind`, `Stage`, `Resource`, `Format`), its five test targets,
`tools/check_iwork_public_api.py` and its tests (they guarded only that
module). The facade `iwork` feature is now the aggregate of `pages`,
`keynote`, `numbers`; `litchi -> litchi-iwa-structured` is gone (that crate
has no workspace dependent left). `litchi::open` no longer rewrites iWork
suggestions; `Error::iwork_unsupported` (only reachable through a caller's
`ZipTail`) names the cargo features.

**Callers moved:**

| Caller | Now |
|---|---|
| Facade tests | `tests/iwork_file.rs` carries the old behaviours over to `File`: directory/file parity (text and outline), a capture that survives removal of its source, loose `Index/` bundles (read, edit, save back as `Index.zip`), typed path failures (missing, unrelated, both index layouts, a conflicting legacy marker, a Unix socket), inclusive directory input limits through `OpenOptions`, the Numbers compatibility fixture, Markdown. `keynote_facade` (immutability of an opened `File` under a leaf edit and a unified set), `pages_facade` (the body storage as `/section[1]`, one `p` per line), `unified_open_facade` (directories open for all three formats), `tolerance_report_facade` (the report of an opened Pages file). |
| Examples, fuzz | `read_iwork` (outline, stats and text of `File`, leaf options set), `profile_all_formats` (Pages/Keynote Markdown, Numbers text), `fuzz_targets/parse_iwork` (open, text, outline, stats, per-kind queries and path round trips, exact no-op publication; `cargo check --bin parse_iwork` 0). |
| litchi-py | `Document`, `Presentation`, `Workbook` hold the iWork `File` and project its leaf package (`File::Pages(h).pages()` etc.): one open instead of a second semantic-only one, the same projections (the Python oracle passes unchanged). `litchi_py.open` opens directories of all three formats; the "open it with Presentation.open" suggestion is gone. New test `test_iwork_package_directories_open_and_save_back_in_place`; the fidelity mirror skips bundle rows (the Rust test owns them). |
| CLI | `iwork_writes_save_in_place` also edits a Keynote bundle in place. |
| Docs | README, `crates/litchi/README.md`, `OFFICE_API_GUIDE.md`, `litchi-iwa` docs, fuzz README, `IWORK_API_REQUESTS.md` (request 16 done), ADR 0024 amendment. |

**Glossary.** New `facade-iwork` profile bans `litchi::iwork`; the
`API_GLOSSARY.md` section "iWork facade retirement" lists the banned names
and the new canonical ones. The gate's rustdoc walker refused the facade's
glob re-exports of other crates (`pub use litchi_pages::*`); a profile may
now name them in `delegated_globs` (here `litchi_core`, `litchi_keynote`,
`litchi_pages`, `litchi_sheet`), each checked by its own profile; any other
glob without module data is still refused (two new gate unit tests).

## 4. Markdown and text through `File`

`File::text()` already read iWork. `ToMarkdown for File` refused Pages and
Keynote as *workbooks*; it now converts them over the object tree
(`litchi-ops/src/markdown/iwork.rs`): Pages body paragraphs, escaped, one
Markdown paragraph each (a document with body tables is refused with
`unsupported_element`, because their place among the paragraphs is not
modelled yet — gap G3); Keynote one `# Slide N <title>` section per slide
with its body, separated by `---`, the presentation family's layout;
front matter from the leaf metadata with `with_metadata`. Numbers stays a
workbook refusal. The Markdown corpus golden covers no iWork input, so no
row changed (`markdown_corpus_golden` passes unchanged).

## 5. Performance

`crates/litchi-ops/examples/iwork_directory_bench.rs` (new), release,
`taskset -c 21`, ABBA × 8 rounds × 30 calls, two runs, host load 6.9–8.4 of
32 cores. Binary SHA-256
`294f21a9ad6aebe6ccc186ea399b84610a093ee2469e8f7c8ddd2b5be6d157fb`.
Evidence: [`results/0306/`](results/0306/) (summary JSONL and both raw runs,
42 KB each).

| Scenario (leaf leg) | Pages | Keynote | Numbers |
|---|---|---|---|
| `dir_open` (`Package::open(dir)`) | 975 → 979 µs, 1.003 / 1.006 | 1866 → 1870 µs, 1.001 / 1.004 | 1171 → 1177 µs, 1.005 / 1.007 |
| `file_open` (`Package::open(file)`) | 0.998 / 1.005 | 0.991 / 0.998 | 1.000 / 1.006 |
| `bytes_open` (`from_shared_bytes_with_options`, no copy) | 1.001 / 1.000 | **1.052 / 1.053** | 0.992 / 1.005 |
| `bytes_open_options` (same leaf leg) | 1.001 / 1.001 | 1.054 / 1.053 | 0.991 / 1.004 |

- The first directory route (classify with `litchi_iwa_detect::path`, then
  the leaf captured again) measured 1.051–1.086 (+78–99 µs); the
  single-capture route of `93f86b8b72` removed it.
- `bytes_open_options` equals `bytes_open` within 0.3 %: the option
  plumbing adds nothing to the default route.
- **Miss (not introduced here):** Keynote `bytes_open` is 5.2–5.5 %
  (+80 µs) over a leaf leg that adopts shared bytes without copying; the
  unified leg moves an owned `Vec` into an `Arc<[u8]>` (a copy of the
  500 KB fixture) and runs detection. Records 0296/0298 measured the same
  route at 0.97–0.98 against `Package::from_bytes`, which copies. Not
  attributed further; follow-up G4.

## 6. Bench matrix

`tools/unified_bench_matrix.py` gained parsers for **0298** (keynote_bench
summary, 10 rows), **0301** (Numbers structure bench, 6), **0302** (the
after legs of its DOCX, ODT, Pages, Numbers and RTF harnesses, 52),
**0303** (pptx_gaps_bench: unified/native after, after/before on the plain
paths, 12), **0304** (text_set_bench: after/before and ODT unified/leaf,
10) and **0306** (§5, 12); `abba_builds` skips a scenario one build refused.
Range 0285–0306; **0305 is listed as "not on this tree"** (its evidence is
not merged). New annotations cite each record's explanation for its open
misses (0298 open + get, 0302 DOCX single set / Numbers open + get / RTF
single-target fan-out, 0304 ODT plain set, 0306 Keynote bytes open).
Regenerated: 1522 rows; `--check` 0.

## 7. Gaps left (follow-ups)

1. **G1 — loose `Index/` bundles save in the `Index.zip` layout.** The IWA
   member bytes are kept; the bundle's physical layout changes on the first
   save.
2. **G2 — Numbers tree on a non-native merge source.** The synthetic
   compatibility fixture's text reads, but `query`/`get` refuse it
   (`invalid_input` from the table-merge reader); a tolerance request for
   the Numbers owner (`IWORK_API_REQUESTS.md`).
3. **G3 — Pages Markdown with body tables** is refused; needs the anchor
   order of tables among paragraphs in the Pages tree.
4. **G4 — Keynote in-memory open copies the owned input** into an `Arc`
   (§5).
5. `litchi-iwa-structured` has no dependent; retire it with the iWork owner.
6. `tools/test_check_crate_boundaries.py`
   `test_iwa_shared_image_adjustments_boundary_accepts_focused_owners` fails
   on the base `e7fe30fb24` too (checked on an archive of the base); not
   touched here.

## 8. Verification

Final tree (after the trunk merge). `CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`, one cargo command at a time; every exit status
checked.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-archive -p litchi-pages -p litchi-keynote -p litchi-numbers -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features pages / keynote / numbers / pages,keynote / pages,numbers / keynote,numbers / markdown,pages / markdown,keynote / markdown,yaml,pages / docx] --all-targets -- -D warnings` | 0 (each, and with no features) |
| `cargo test -p litchi-iwa-archive --all-features` | 0 (200) |
| `cargo test -p litchi-pages --all-features` | 0 (684) |
| `cargo test -p litchi-keynote --all-features` | 0 (1,144) |
| `cargo test -p litchi-numbers --all-features` | 0 (1,057) |
| `cargo test -p litchi-ops --all-features` (`--lib`, `--doc`, then the integration tests in groups of 40, executables pruned between groups) | 0 (2,341) |
| `cargo test -p litchi --all-features` | 0 (230) |
| `cargo test -p litchi-cli --all-features` | 0 (33) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa-archive -p litchi-pages -p litchi-keynote -p litchi-numbers -p litchi-ops -p litchi -p litchi-py -p litchi-cli` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (278) |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, then with every feature but `encryption` | 0, 0; goldens unchanged (`git diff --exit-code` 0) |
| `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI` | 0 (table unchanged) |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test corpus_fidelity` (after the trunk merge) | 0 (re-blessed: + the three bundles) |
| `python3 tools/check_crate_boundaries.py` | 0 (67 packages, 282 declarations, 11 debts) |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `python3 -m unittest tools.test_non_iwork_gate tools.test_check_api_glossary` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile facade-iwork / ops-iwork / pages / keynote-minimal / keynote-full / numbers-minimal / numbers-full / ops-odf / ops-rtf` | 0 (each) |
| `python3 tools/check_iwork_leaf_public_api.py` | 0 |
| `python3 tools/unified_bench_matrix.py --check` | 0 |
| `cd crates/litchi/fuzz && cargo check --bin parse_iwork` | 0 |

The OfficeCLI corpus has no iWork input and no other behaviour changed, so
it was not re-run.
