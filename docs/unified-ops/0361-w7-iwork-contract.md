# 0361 — W7: iWork on the record 0351 contract

- Status: implemented on `uo2/w7-iwork`, cut from `feat/unified-ops-api` at
  `7f36f7b909`, merged with trunk `f8899dd2cb` (records 0356b, 0357b, 0358,
  0358a) and `720390ef28` (record 0359).
- Date: 2026-10-20.
- Input: owner decision [0350](0350-owner-api-debt-decision.md) items
  14–19; contract [0351](0351-api-contract.md) §2, §4.3, §5 (wave W7);
  record [0352](0352-iwork-names.md) §7 (the `litchi-iwa` leftovers).
- Owned: `crates/litchi-{pages,keynote,numbers}/**` and their fuzz crates;
  `crates/litchi-iwa/**`; the iWork hunks of `crates/litchi-ops/src/error.rs`
  (`refusal_code`), `crates/litchi-ops/src/format/iwork/**`, the iWork
  benches and tests of litchi-ops, the facade's iWork tests and fuzz
  targets; the iWork profiles of `tools/api_glossary.json`; the iWork audits
  of `tools/check_crate_boundaries.py` and its tests;
  `tools/check_iwork_leaf_public_api.py` and its tests; the W7 line and the
  *Split commit* cell of `docs/API_GLOSSARY.md`; current-state iWork docs.
- `performance_claim`: measured (§8).

## 1. Summary

| 0351 rule | Before | After |
|---|---|---|
| §2.2 commit triad | 93 scoped commits without `into_parts`; `keynote::slide::delete::Patch` without `is_noop` | every iWork commit has `package`, `into_package`, `patch`, `diagnostics`, `into_parts`; `family_methods: true` on `pages`, `numbers-minimal/full`, `keynote-minimal/full` |
| §2.3 every family registered | six families reachable only through one-shot `Package` verbs (no `Edit`); `section::paragraphs` unregistered | each has an `Edit` staging one operation and `commit()`; all registered (pages 30, numbers 32, keynote 44 families) |
| §2.3 no aliases | `apply_slide_media_lifecycle_patch`, `physical_sort::TransactionError`, `slide::media::MediaInfo`, `numbers::table::relocation::transaction` | deleted |
| §2.4 error taxonomy | 105 family errors mixing refusal reasons with standard variants; `LimitExceeded` | `Error::Refused(Refusal)` with a typed module `Refusal` and `Refusal::kind()`; `Limit { kind, observed, maximum }`; litchi-ops classifies by kind |
| §2.5 no `*_with_limits` | `pages::MergeReader::{open,from_bytes,from_shared_bytes}_with_limits`, `numbers::Workbook::from_sheets_with_limits`, host `MediaManager::*_with_limits` | `*_with_options(.., OpenOptions)`; `_with_limits` banned in the five leaf profiles and `iwa` |
| §2.1 immutable owners (host) | `IWorkPackage` with public `&mut self` mutators; mutable `Editor` owners | `litchi_iwa::Package` (mutators crate-private), registered immutable with `litchi_iwa::Document`; the editors are `Builder` drafts named `Draft` |
| 0352 §7 deprecated items | 18 `#[deprecated]` host items | none: 16 kept as documented native-ID API, `Borders` alias and the root re-export deleted |
| 0352 §7 `IWork*` prefixes | `IWorkPackage`, `IWorkTextEditor`, `IWorkDrawableCommentEditor`, `IWorkMediaEditor` | `Package`, `text::Draft`, `comments::Draft`, `media::Draft`; `IWork` banned as a prefix |
| tool test | `tools/test_check_iwork_leaf_public_api.py`: 42 errors | 16 pass |

Every iWork round-trip, CRUD, corpus and no-op test passes with
byte-identical output (§7). The OfficeCLI corpus is unchanged.

## 2. Commit triad and families (leaf crates)

`into_parts(self) -> (Package, Patch)` is added to the 93 scoped commits
the 0351 inventory listed (pages 24, numbers 27, keynote 42). `keynote::slide::delete::Patch::is_noop` reports
`false`: a deletion, or its inverse restoration, always changes the slide
count.

Six families had a `Patch`, `Commit` and `Error` but were published by
one-shot `Package` verbs with no `Edit`. Each now has the single-operation
shape of `pages::section::lifecycle` (record 0312): `Package::edit_x()`
returns an `Edit` holding the package borrow; a consuming verb stages one
operation; a second verb is `OperationAlreadyStaged`, an empty commit
`NoStagedOperation`; `commit()` runs the former verb's body unchanged.

| Family | Entry | Edit verbs | Replaced |
|---|---|---|---|
| `pages::table::creation` | `edit_body_table_creation()` | `insert(anchor, NewTable)` | `insert_body_table` |
| `pages::table::deletion` | `edit_body_table_deletion()` | `remove(selector)` | `remove_body_table` |
| `keynote::slide::media::lifecycle` | `edit_slide_media_lifecycle()` | `duplicate`, `remove`, `duplicate_movie`, `duplicate_audio`, `remove_movie`, `remove_audio` | the six `*_slide_media/movie/audio` verbs |
| `keynote::slide::audio::creation` | `edit_slide_audio_creation()` | `add(slide, filename, data, options)` | `add_slide_audio` |
| `keynote::slide::movie::creation` | `edit_slide_movie_creation()` | `add(slide, movie name, movie, poster name, poster, options)` | `add_slide_movie` |
| `pages::section::paragraphs` | (already had `Edit`) | — | registration only |

The staged inputs are borrowed (`&'a [u8]`, selectors with `'a`), so an edit
copies no payload; the commit does exactly the work the verb did. The
pages creation `Diagnostics` gains `changed()`. Every capability of the
removed verbs is reachable through the new verbs (the kind-checked
duplicate/remove routes stay distinct verbs).

Aliases deleted: `Package::apply_slide_media_lifecycle_patch` (a
"compatibility spelling" of `apply_slide_media_lifecycle`),
`keynote::slide::table::physical_sort::TransactionError` (`= Error`),
`keynote::slide::media::MediaInfo` (`= MovieInfo`) and the
`numbers::table::relocation::transaction` namespace (a second path to
`table::relocation`).

## 3. Error taxonomy (0351 §2.4)

The 105 family `Error` enums (pages 30, keynote 44, numbers 31) were each
split:

- `Error` keeps the standard variants: `Limit { kind, observed, maximum }`
  (renamed from `LimitExceeded`), `PatchConflict`, `InvalidSource`,
  `Allocation`, `Verification`, and the wrappers of inner errors (`Text`,
  `Values`, …), plus `Refused(Refusal)` (`#[error(transparent)]`);
- every refusal reason (`*NotFound`, `Ambiguous*`, `Empty*Name`,
  `Unsupported*`, `TableLocked`, `WrongFormatFamily`, `CannotDeleteFinalSlide`,
  `OperationAlreadyStaged`, …) moves, with its name, payload, doc and
  message, into the module's typed `Refusal`, re-exported from the family's
  public module;
- `Refusal::kind() -> litchi_core::edit::RefusalKind`; `impl From<Refusal>
  for Error`.

Display output is unchanged (transparent), `Debug` shows `Refused(..)`.
`numbers::table::cells::Error` (the hot cell path) had a hand-written
`Display` and was split by hand with the same messages. The reasons are
content-free as before (positions and counts only).

`LimitExceeded` is also `Limit` in the public read-side errors that used it
(`pages::table::catalog`, `pages::chart::metadata`, `pages::table::cells`,
`numbers::chart::metadata`, `numbers::ResourceError`, `numbers::formula`,
`keynote::chart::metadata`, `keynote::slide::table::cells`), and
`numbers::table::edit::Error::Limit` gains a typed `LimitKind::Updates`
(it carried an untyped `requested`/`maximum` pair). `LimitExceeded` is
banned in the leaf profiles.

Kinds: `NotFound` for missing targets; `InvalidValue` for ambiguous or
out-of-range selectors and invalid values (and the Keynote layout
selectors, which name a property value); `Protected` for locked tables;
`UnsupportedTarget` for a target of the wrong kind or a final element;
`UnsupportedOperation` for staging misuse and styling an empty cell;
`UnsupportedContent` for sources and dependencies the route cannot rewrite.

litchi-ops: `crate::error::refusal_code(kind)` is the one kind→code table
(the F0 `Classify for RefusalKind` now uses it); every iWork leaf
classifier maps `Refused(refusal)` by `refusal.kind()` alone, `Limit` to
`limit_exceeded`, `InvalidSource`/`PatchConflict` to `invalid_input`, the
rest to `internal_error`. One exception, the PPTX precedent
(`SlideRemovalRefusal::FinalSlide`): removing the final slide, section,
paragraph, sheet or table axis stays `invalid_argument`. Codes that changed
for a leaf refusal (the unified layer resolves paths itself first, so these
are reached only when the leaf refuses): `path_not_found` → `not_found`,
`invalid_path`/`out_of_range`/`invalid_argument` → `invalid_value`,
wrong-kind targets `unsupported_capability` → `unsupported_type`. No
OfficeCLI corpus, capability or registry output changed.

## 4. Configured limits (0351 §2.5)

| Before | After |
|---|---|
| `pages::MergeReader::{open, from_bytes, from_shared_bytes}_with_limits(.., litchi_iwa_archive::Limits)` | `{open, from_bytes, from_shared_bytes}_with_options(.., pages::OpenOptions)` (its `source_limits`; the keynote/numbers readers' shape) |
| `numbers::Workbook::from_sheets_with_limits(sheets, SemanticLimits)` | `from_sheets_with_options(sheets, OpenOptions)` (its `semantic_limits`); the limits form is crate-private |
| host `MediaManager::{new, new_with_limits, from_bytes_with_limits, from_package_with_limits}` | `open`, `from_bytes`, `from_package` and `*_with_options(.., media::OpenOptions)` (`media_limits`, `media_limits_value()`) |
| host `IWorkPackage::*_with_limits` (public but on an unnameable type) | crate-private |

The iWork `OpenOptions` getters were already `*_value()`; no iWork crate has
an `EditOptions`.

## 5. `litchi-iwa` (0352 §7)

### 5.1 Owner and drafts

`litchi-iwa` is the unpublished migration host (record 0352 §2.1: no
workspace manifest depends on it; ADR 0028 keeps it until the leaf crates
reach parity). Its types are classified per item:

| Item | Decision | Why |
|---|---|---|
| `IWorkPackage` → `litchi_iwa::Package` | owner, immutable | the host's archive value, now nameable; its four `&mut self` mutators (`insert_entry`, `remove_entry`, `replace_archive`, `update_archive`) and four `*_with_limits` constructors had no caller outside the crate and are crate-private; registered in `immutable_owners` |
| `litchi_iwa::Document` | owner, immutable | already conforming; registered |
| `pages::Editor`, `keynote::Editor`, `numbers::Editor` → `pages::Draft`, `keynote::Draft`, `numbers::Draft` | draft (0351 §2.1 rule 7) | `Builder::build()` returns it; it is a mutable working copy that verifies after each change and publishes only through `to_bytes`/`save`/`into_package` — no patch, commit or diagnostics, so it is not a published owner. Converting its ~1,900 native-ID methods into package-level families would duplicate the leaf `Package` families the host is being retired into. It lives in the per-format module next to `Builder`; the `editor` modules are crate-private (their 16 items with no other public path, e.g. `TableCellInset`, are re-exported from the format modules) |
| `IWorkTextEditor` → `text::Draft` | draft | same shape, over the shared TSWP text layer |
| `IWorkDrawableCommentEditor` → `comments::Draft` | draft | same shape; root re-export deleted |
| `IWorkMediaEditor` → `media::Draft` | draft | same shape |

Crate-private names keep their prefix (`IWorkChartArchive`,
`IWorkThemeArchive`, …): they are not public API. Inside the crate the host
package is imported as `HostPackage` (a private alias) so it never shadows
the leaf `Package` types those modules import.

### 5.2 Deprecated items (18)

| Item | Decision | Parity / reason |
|---|---|---|
| `numbers::editor::table::cell::Borders` re-export | deleted | `litchi_iwa_common::table::cell::Borders` is the same type |
| root `IWorkDrawableCommentEditor` re-export | deleted | `comments::Draft` |
| `IWorkDrawableCommentEditor` (struct) | kept as `comments::Draft` | Pages and Numbers drawable comments have no leaf owner |
| `pages::Draft::{drawable_comment, set_drawable_comment, clear_drawable_comment}` | kept, plain | no `litchi-pages` comment owner |
| `numbers::Draft::{sheet_drawable_comment, set_sheet_drawable_comment, clear_sheet_drawable_comment, sheet_drawable_comment_replies, add/set/remove_sheet_drawable_comment_reply}` | kept, plain | no `litchi-numbers` sheet-drawable comment owner |
| `numbers::Draft::{set_cell_comment, cell_comment_replies, add/set/remove_cell_comment_reply}` | kept, plain | strict graphs already delegate to `litchi_numbers::Package::edit_table_cell_comment*`; the fallback for graphs that owner refuses (comment-storage creation, reply identities) has no leaf equivalent |

The kept items carry a "Native-ID host API (record 0361)" doc note. The 73
`#[allow/expect(deprecated)]` suppressions that only covered them are
removed; 43 that cover deprecated protobuf fields stay. The boundary audits
that required these items to stay deprecated now require the opposite (no
deprecation, the native-ID note).

### 5.3 Gate

Profile `iwa`: bans `IWork` (prefix), `Editor` (exact and suffix) and
`_with_limits` (suffix); `immutable_owners`: `litchi_iwa::Package`,
`litchi_iwa::Document`. The value enum `numbers::UnverifiedTopologyEdit`
(the 0351 inventory's heuristic row) is `UnverifiedAxisChange`.

## 6. Tools

- `tools/check_iwork_leaf_public_api.py` checks each leaf at its own family
  projection module again (`semantic_module`: numbers → `workbook`, pages →
  `document`, keynote → `presentation`); `tools/test_check_iwork_leaf_public_api.py`
  (16 tests) passes. It is not yet in the coordinator's tool unit-test
  list; the coordinator should add it.
- `tools/check_crate_boundaries.py`: the deletion, media-lifecycle,
  audio/movie-creation and merge-reader audits follow the `Edit` routes and
  `*_with_options`; family exports may include `Refusal`; the host audits
  follow the renamed types and the inverted deprecation rule. 1,139 unit
  tests pass.
- `tools/api_glossary.json`: §1 registrations.

## 7. Fidelity

- Leaf suites (pages, keynote, numbers, numbers-wire, iwa-common) and
  litchi-iwa: 5,126 tests, 0 failures, including the exact-bytes
  round-trip, inverse-restores-source and no-op tests of every family.
- litchi-ops: `--lib` and the iWork, corpus-fidelity (whole-corpus
  open → empty edit → save byte-identical, the 126 iWork no-op saves
  included), batch-equivalence, capabilities and registry targets: 1,040
  passed, 0 failed; then the full suite in chunks (§9).
- New tests: single-operation staging (`OperationAlreadyStaged`,
  `NoStagedOperation`, discard leaves the source) for table creation,
  media lifecycle and audio creation; `is_noop` of a deletion patch.

## 8. Performance

Method (record 0329's, the 0351 §5 W7 routes): release builds of the three
in-process paired benches (`cargo build --release -p litchi-ops --features
<format> --example {keynote_bench, pages_bench, numbers_unified_bench}`, one
format feature each) and of record 0337's `unified-batch-closure` harness
(features `pages,keynote,numbers`), from trunk `f8899dd2cb` ("base") and
this branch after the trunk merge ("after"). One pinned CPU (`taskset -c
30`) of the shared 32-core host at load 5.3–12.9; processes B A A B; each
bench interleaves leaf and unified samples in-process (keynote/pages 8
rounds × 30 iterations, numbers 8 samples, batch 5-round ABBA). Ratio =
unified / leaf. Evidence: [results/0361/](results/0361/) (`build.sh`,
`measure.sh`, `summarize.py` → `summary.txt`, raw JSON, `abba/`,
`binaries.sha256`, `callgrind-keynote-10k.txt`; 572 KB). The numbers bench
includes record 0347's sheet/table swap scenarios; every numbers scenario
reported identical outputs, and the batch harness checks exact bytes before
timing.

| Route | Budget | base (procs 1, 4) | after (procs 2, 3) |
|---|---|---|---|
| Keynote open | 1.05 / +20 µs | 0.992–0.994 | 0.991–0.994 |
| Keynote set (slide text) | 1.10 | 1.004–1.007 | 1.010–1.011 |
| Keynote `File::batch` 1k (bench) | 2.0 | 1.047–1.056 | 1.044 |
| Keynote structural (add/remove slide, duplicate, media remove, table add) | 1.10 | 0.992–1.030 | 0.992–1.029 |
| Pages open | 1.05 / +20 µs | 0.991–0.993 | 0.995–0.997 |
| Pages set / add / remove / move | 1.10 | 1.008–1.019 | 1.008–1.019 |
| Pages batch 1k / 10k (bench) | 2.0 / scaling | 1.067–1.069 / 1.417–1.432 | 1.069 / 1.409–1.432 |
| Numbers open + get | 1.05 / +20 µs | 1.066–1.086 (+6.6–8.5 µs) | 1.069–1.074 (+6.9–7.2 µs) |
| Numbers set | 1.10 | 0.991–1.020 | 0.989–1.009 |
| Numbers batch 1k (bench) | 2.0 | 1.300–1.310 | 1.269–1.298 |
| Numbers sheet/table swap (0347) | 1.10 | 0.991–1.003 | 0.989–1.014 |
| batch-closure keynote 1k / 10k | 2.0 | 1.094–1.098 / 1.720–1.745 | 1.104–1.109 / 1.805–1.848 |
| batch-closure pages 1k / 10k | 2.0 | 1.068–1.083 / 1.583–1.598 | 1.076–1.078 / 1.555–1.567 |
| batch-closure numbers 1k / 10k | 2.0 | 1.340–1.370 / 1.477–1.510 | 1.352–1.360 / 1.453–1.464 |

Every route meets its budget (the Numbers open is above 1.05 but within
the +20 µs allowance, as at base). Leaf times are unchanged between base
and after (Keynote set 4,981–5,023 / 4,945–4,970 µs, Pages set
2,006–2,015 / 2,004–2,011 µs, Numbers set 369–379 / 371–377 µs): the
single-operation `Edit`s borrow their inputs and run the former verb, and
the error split changes only the error path.

The one movement is the Keynote 10k batch-closure case (1.72–1.75 →
1.81–1.85; unified 10.19–10.33 → 10.67–10.84 ms; leaf 5.92 ms both). Its
instruction count is unchanged: callgrind of one process of three 10k
rounds is 2,042,624,058 Ir at base and 2,041,135,054 after (−0.07 %), so
the wall-time difference is placement/load noise, not added work. It stays
within the 2.0 budget. The bench matrix gains 65 rows for record 0361, all
meeting their budgets.

## 9. Verification

On the tree merged with trunk `f8899dd2cb`; every command exited 0.

| Gate | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-{pages,keynote,numbers,numbers-wire,iwa-common,iwa} -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features pages\|keynote\|numbers] --all-targets -- -D warnings` | 0 (all four) |
| `cargo test --all-features -p litchi-pages -p litchi-keynote -p litchi-numbers -p litchi-numbers-wire -p litchi-iwa-common -p litchi-iwa` | 0; 5,126 passed, 0 failed |
| `cargo test -p litchi-ops --all-features`, chunked (`--lib`, `--doc`, 9 groups of ≤20 `--test` targets, executables deleted between chunks) | 0 every chunk; 2,642 passed, 0 failed, 7 ignored |
| `cargo test -p litchi --all-features`; `-p litchi-cli` (OfficeCLI corpus included) | 0; 243 passed |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0; 301 tests |
| capability/registry goldens, all features and all-but-`encryption` | 0 both; unchanged, nothing re-blessed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-{pages,keynote,numbers,iwa} -p litchi-ops -p litchi` | 0 |
| `cargo check --bins` in the `litchi`, `litchi-ops`, `litchi-pages`, `litchi-keynote`, `litchi-numbers` fuzz crates | 0 (the pre-existing dead-code warnings only) |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0: 44 profiles, 3,113 banned names, 461 families, 73 immutable owners, family methods in 27 profiles |
| `python3 tools/check_iwork_leaf_public_api.py` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0, 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_check_iwork_leaf_public_api tools.test_check_api_glossary tools.test_unified_bench_matrix` | 0; 1,320 tests |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py` (31 ok) | 0 each |

`tools/perf-baseline/unified-batch-closure/Cargo.lock` was stale on trunk
(it lacked `litchi-sheet-text` and `quick-xml` for `litchi-ops`, so
`cargo build --locked` of the harness failed); it is refreshed with
`cargo metadata --offline` (12 added lines, no version change).

### 9.1 After the trunk merge `720390ef28` (record 0359)

Merge `cee98f7b28` (conflicts only in the bench-matrix tool's parser list
and its generated Markdown, regenerated). Re-run, each exit 0: fmt; the
all-features clippy above and the four litchi-ops feature sets; leaf,
litchi-iwa, facade and cli tests (5,369 passed); the chunked litchi-ops
suite (2,642 passed, 7 ignored); goldens in both feature sets (unchanged);
pyo3 `run_tests.sh`; docs; perf-baseline `--locked`; the full glossary
gate (47 profiles, 3,207 bans, 492 families, 73 immutable owners, family
methods in 30 profiles); boundaries; non-iWork gate; the tool unit tests;
the OfficeCLI delta/corpus, bench-matrix and feature-matrix (31 ok)
checks.

### 9.2 Commits

| Commit | Content |
|---|---|
| `9810d38be5` | leaf triad, single-operation families, options ingress, aliases, gate registration |
| `23f46568c6` | `check_iwork_leaf_public_api` per-crate semantic module |
| `e0632a4240` | `litchi-iwa`: `Package`, drafts, no deprecated items, `media::OpenOptions` |
| `a587b2b3d5` | refusal taxonomy, `Limit`, litchi-ops classification by kind |
| `624c8092ba` | classification test, `UnverifiedAxisChange`, harness lock |
| `5a960f58f1` | this record, evidence, bench matrix, glossary W7 line |

## 10. Gaps and follow-ups

- `litchi-iwa-common::wire` keeps limit-less/limited twins
  (`parse_wire_fields{,_with_limits}`, `WireView::parse{,_with_limits}`,
  `preflight_wire_tree*`, `patch_nested_fields_batched*`,
  `append_length_delimited_field*`). They are codec functions not exposed
  through any leaf crate's public API; 0351 §2.5 rule 3 wants one form with
  a required limits argument (about 2,000 call sites across the iWork
  crates). Follow-up F-0361-WIRE.
- `litchi-iwa-structured::{Snapshot::from_parts_with_limits,
  from_keynote_presentation_with_limits, from_pages_document_with_limits}`:
  no workspace consumer and no gate profile. Follow-up F-0361-STRUCTURED
  (move to `*_with_options` or crate-private).
- The host drafts keep `&mut self` by design (§5.1); retiring them is ADR
  0028's deletion gate, not this wave.
