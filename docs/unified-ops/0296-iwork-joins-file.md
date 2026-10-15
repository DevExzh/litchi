# 0296 — iWork joins `File`: Pages, Keynote and Numbers (M6 foundation)

- Status: implemented on `uo2/iwork-found`, cut from `feat/unified-ops-api`
  at `943be1a927`. Plan [0248](0248-coordinator-restart-plan.md) §6 Phase F,
  step 24 (foundation), and step 28 (partial, §5).
- Date: 2026-10-15
- Owner decisions applied: **D1** (remove the iWork-free gate; `File` gets
  `Pages`, `Keynote`, `Numbers` directly) and **D2** (the iWork crates may be
  edited), both of 2026-10-15 (`PROGRESS.md`, "Owner decisions").
- Owned: `crates/litchi-ops/src/format/iwork*` (new), small separated hunks
  in the shared ops files, the gate tools, the facade's open path, one
  additive hook in `litchi-numbers`.
- `performance_claim`: paired leaf/unified medians in §6 (release, one
  pinned core of a shared host, ABBA). No frozen-harness qualification.

Before this record `litchi::open` recognized iWork packages (through the
facade's ZIP tail) and refused them with `unsupported_format`; `litchi-ops`
was forbidden any iWork dependency (`tools/non_iwork_gate.py`), and the
capability reports were whole-format `M6` refusals. Every claimed format
now opens into a `File` variant.

## 1. Gate removal (D1)

- `crates/litchi-ops/Cargo.toml`: `pages`, `keynote` and `numbers` gain
  their `dep:` edges (`litchi-pages`, `litchi-keynote`, `litchi-numbers`,
  each with its leaf's `internal-iwork-source` prepared-source ingress) and
  turn on the internal umbrella `iwork-adapter` (`dep:litchi-iwa-detect`).
- `tools/crate_boundaries.json`: four new `litchi-ops` edges; the reason
  ("owner decision D1, record 0296") is recorded in the ADR 0024 amendment
  of 2026-10-15 (the JSON has no per-edge reason field). ADR 0033 §7's
  planned reclassification is marked done.
- `tools/non_iwork_gate.py`: `litchi-ops` and `litchi-cli` are derived as
  the *iWork-capable unified layer* (`unified_packages`, checked against a
  reviewed inventory): excluded from the bulk lane (48 → 46 bulk, 18 → 20
  excluded) but allowed in the safe facade trees, whose forbidden set is now
  `excluded − unified`. The gate stays the fast non-iWork lane for the leaf
  and common crates, and still proves the safe facade closures reach no
  iWork package. Any *other* package that grows an iWork edge still fails
  the bulk boundary with the offending edge. `tools/test_non_iwork_gate.py`
  models the new edges and asserts the partition.
- `tools/validate_crud_coverage_index.py` `FORBIDDEN_IWORK` is unchanged:
  it scopes the performance program's non-iWork CRUD selector registry
  (`tools/perf-baseline`), not the unified layer, so it does not block D1.

## 2. `File::{Pages, Keynote, Numbers}`

One object-tree adapter (`litchi_ops::format::iwork`, handle `Iwork`) in
the shape of the ODF object-tree adapter of record 0282: the handle holds
the leaf `Package`; a tree of items (kind, properties, text, leaf
positions) is built from the leaf's read-only projection **on the first
read**, so `open` decodes only what the leaf open decodes; one generic
`Nav<F: Flavor>` serves `get`/`find`/`query`/`view`; writes run the leaf's
focused transactions in order and adopt the last published package.

| Format | Elements (paths) | Writes (`set`) | Other verbs |
|---|---|---|---|
| Pages | `document` `/`; `body` `/body`; `p` `/body/p[N]` (one per line of a section's text, `section` prop); `section` `/section[N]` (`name`, `text`, `paragraphCount`) | `p.text` (UTF-16 span of its line in the section, via `edit_section_text`); `section.text` (`set_section_text`) | `create`, `add`, `remove` refused (typed gaps, requests 1, 11, 12) |
| Keynote | `presentation`; `slide` `/slide[N]` (`name`, `title`, `hidden`); `placeholder` `/slide[N]/placeholder[title\|body]` (`type`, `text`); `notes` | `placeholder.text` (`edit_slide_text`); `slide.hidden` (`edit().set_slide_skipped`) | `remove` slide (`edit_slide_deletion`), `move` slide (`edit_slide_order`); `add`, `swap`, `create` refused |
| Numbers | `workbook`; `sheet` `/sheet[N]` or `/<name>`; `table` `/sheet[N]/table[M]` or `/<sheet>/<table>`; `cell` `/…/table[M]/A1` (`type`, `value`, `formula`) | `cell.value` (text, number, boolean, `null` clears; `edit_table_cells`, one leaf transaction per table); empty cells inside the grid are addressable for `set` | `add`, `remove`, `move`, `create` refused |

- **Detection and open.** `litchi-ops` classifies iWork itself, as the last
  built-in ZIP candidate (after OOXML and ODF, before a caller's tail):
  `PreparedSource::from_shared_bytes` prepares the container once and the
  leaf adopts it (`__from_prepared_source`; record 0296 adds the Numbers
  ingress, commit `b93ddf8825`). A detector refusal declines the input
  (reported unrecognized, as before). A format whose feature is off is
  `unsupported_format` with its feature. Path opens attach the origin, so
  `save()` replaces the file atomically.
- **Package directories** (the app-authored bundle form) are refused with
  `Error::iwork_unsupported` (now: "the unified `File` cannot hold it"):
  the leaves refuse directory sources for editing ("semantic projection
  only"). The facade's read-only `litchi::iwork::Document` reads them.
- **Save / publish.** `to_bytes`/`write_to` stream the leaf's exact bytes;
  `save`/`save_as` stage them through `litchi_core::atomic::save`
  (sibling temporary, durability and overwrite options). A no-op publish is
  byte-identical (corpus fidelity, §4).
- **Metadata.** `File::metadata()` returns the leaf's metadata (Pages,
  Keynote properties; Numbers document metadata). `/metadata` is not a path
  (gap; request 14 for Pages writes).
- **Errors.** Leaf errors stay typed behind `Source::Iwork(LeafError)` (one
  variant per leaf error type) and map to the closed code list
  (`LeafError::code`): limits → `limit_exceeded`, selector misses →
  `path_not_found`, unsupported native shapes → `unsupported_capability`,
  locked Numbers tables → `document_protected`.
- **Capabilities.** `capabilities()` and `schema()` report the three spec
  tables, the writable properties, `IWORK_STRUCTURE` (follow-up
  `IWORK_API_REQUESTS.md`) for the structural verbs the leaves lack, and
  `IWORK_CREATE` (request 1) for `create`. `ToleranceReport::Untracked`.
- **Not done here.** No batch coalescing (each `set` is its own leaf
  transaction chain; `batch`/`edit` work through the generic executor);
  no `Locked` iWork state (request 3).

## 3. Facade, CLI, Python

- `litchi::open`/`open_bytes` call `litchi_ops::open_with`/`open_bytes_with`
  directly; the facade keeps only the suggestion rewrite for builds without
  the `iwork` feature. The facade's detection functions still use its
  `IworkTail` (they only classify). `crates/litchi/src/lib.rs` docs updated.
- `litchi-cli` enables `pages`, `keynote`, `numbers`; `tests/cli.rs`
  `iwork_writes_save_in_place` sets a Pages paragraph, a Keynote title and a
  Numbers cell in place and reads them back; `create blank.pages` exits 5.
- `litchi-py`: `litchi_py.open` returns iWork `File`s; the family classes
  (`Document`, `Presentation`, `Workbook`) keep the semantic readers for
  iWork (the family views do not cover iWork). Tests:
  `test_iwork_packages_open_as_files`,
  `test_iwork_package_directory_refusal_names_the_family_opener`,
  `test_iwork_writes_round_trip`; docstrings/stub/README updated.

## 4. Tests and goldens

- `crates/litchi-ops/tests/iwork_unified.rs` (13 tests): open routes, reads,
  queries, views, paragraph/section/title/body/hidden/cell writes with
  reopen, exact no-op publish, typed refusals, atomic path save.
- Facade: `unified_open_facade.rs` (iWork opens on every route; directories
  refused; detection agrees with open; Pages tolerance untracked),
  `pages_facade.rs`.
- Query corpus: `PAGES`, `KEYNOTE`, `NUMBERS` selectors over every
  `.pages`/`.key`/`.numbers` file in `test-data/`, recorded from this
  adapter's reads (no earlier iWork digest exists to compare with).
- Phase G checklist (`tests/crud_matrix.rs`) and corpus fidelity
  (`tests/corpus_fidelity.rs`) include the three formats; results in §4.1.
- Capability and registry goldens re-blessed in both feature sets.

### 4.1 Results

- Corpus fidelity (`docs/unified-ops/results/0289/fidelity.{md,tsv}`,
  re-blessed): every iWork file opens and its open → empty edit → save is
  **byte-identical** — Pages 32, Keynote 55, Numbers 39 (126 files; the
  corpus total goes from 771 to 897). No normalization is claimed. The
  set-then-revert probe: Pages 16 `byte-exact` and 16 `refused` (typed
  `unsupported_capability`: "the Pages text edit intersects dependent
  content" — body tables, charts and other anchored objects); Keynote 55
  `value-exact` (a slide-text commit rewrites the package, e.g. drops the
  stale slide preview, so the reverted bytes are smaller than the input); Numbers 39 `none` (cells hold no
  `text`-typed settable property).
- Phase G checklist (`docs/unified-ops/results/0289/crud-checklist.*`,
  re-blessed): Pages 4 elements, 25 claimed cells, 25 passed; Numbers 4,
  22, 22; Keynote 4, 26 claimed, 21 passed, 1 known failure (F8: the corpus
  has only one-slide decks, and the leaf refuses deleting the last slide),
  4 unreached (no fixture has speaker notes). Every other cell is a typed
  gap with its follow-up.
- Query corpus (`tests/golden/query_corpus/digests.json`): 25 new keys
  (31 Pages, 54 Keynote, 38 Numbers files opened, none skipped); every
  existing digest unchanged.
- Capability and registry goldens: new `pages.json`, `keynote.json`,
  `numbers.json` in `tests/golden/{capabilities,registry}`, identical in
  the all-features set and the all-but-`encryption` set; the OfficeCLI
  delta table (`officecli-capability-delta.md`) regenerated (+92 lines,
  iWork sections only).

## 5. `litchi::iwork::Document` (plan step 28): kept

`File` now covers the reads of single-file packages, but the facade reader
stays, for these precise gaps:

1. **Package directories** — the leaves refuse directory sources for
   editing; `File` refuses them (request 16).
2. **Pages body tables and Keynote/Numbers table text roles** — the snapshot
   reports table names and cells inside Pages and Keynote documents; the
   unified Pages and Keynote trees do not model tables yet (0297, 0298).
3. **Text roles and section kinds** — `TextRole`/`SectionKind` (headers,
   footers, footnotes, table names) have no unified element yet.
4. **The `Options`/`SnapshotLimits` profile** — the unified open has no
   iWork sub-options in `OpenOptions` yet (the leaves' defaults apply).

Python's family classes depend on the per-format semantic readers, not on
`litchi::iwork::Document`. Delete it when 1–3 close.

## 6. Performance

`crates/litchi-ops/examples/iwork_bench.rs`: leaf vs unified over the same
bytes, one process, release build, pinned to core 31 (`taskset -c 31`) of a
shared 32-core host (load ≈ 7), ABBA × 8 rounds × 30 calls, three runs.
`open` = in-memory open (leaf `Package::from_bytes` vs `open_bytes`, owned
bytes prepared outside the timed region); `open_get` = open + one element's
text; `set` = one text/cell write and commit on an open package (unified
includes building the object tree on the first read). Per-run median ratio
unified/leaf (summary: `docs/unified-ops/results/0296/iwork-bench-summary.jsonl`;
raw JSONL kept outside the repository, sha256 `abf9da20…`, `4b492f4d…`,
`7f22409e…`, regenerate with the command in the example's header):

| Format | open | open + get | set |
|---|---|---|---|
| Pages (931–946 µs open) | 1.002 / 1.000 / 1.009 | 1.004 / 1.002 / 1.011 | 1.003 / 1.002 / 1.002 |
| Keynote (1.52–1.64 ms open) | 0.972 / 0.966 / 0.997 | 0.985 / 0.980 / 1.011 | 1.005 / 1.003 / 1.005 |
| Numbers (1.05–1.06 ms open) | 1.000 / 0.995 / 1.000 | 1.002 / 0.996 / 1.003 | 1.004 / 1.005 / 1.004 |

Every pair is inside the budgets (open ≤ 5 % or 20 µs; single-op set
≤ 10 %): the largest unified excess is 1.1 % (≈ 11–15 µs) on `open_get`.
The unified open is detection (`PreparedSource`) plus the leaf's
prepared-source ingress, which replaces the leaf's own container parse, so
no second ZIP pass is paid. No regression to profile, so no callgrind run
was made. Not measured: batches (no iWork coalescing yet), large documents
(the corpus fixtures are small; the tree build is linear in stored
cells/paragraphs and runs once per handle).

## 7. Verification

All with `CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`, exit status 0 unless
stated:

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-ops -p litchi -p litchi-py -p litchi-cli
  --all-targets --all-features -- -D warnings`; `cargo clippy -p
  litchi-numbers --lib --all-features -- -D warnings` (its test targets
  carry pre-existing lints from the iWork waves — `needless_update`,
  `manual_contains`, `items_after_test_module` — untouched here);
  `cargo clippy -p litchi-ops --no-default-features [--features F]
  --all-targets -- -D warnings` for no features, `pages`, `keynote`,
  `numbers`, `pages,keynote`, `pages,numbers`, `keynote,numbers`, `docx`.
- `cargo test -p litchi-ops --all-features` (2209 passed),
  `cargo test -p litchi --all-features` (260), `cargo test -p litchi-cli
  --all-features` (32), `cargo test -p litchi-numbers --all-features`
  (1034).
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops
  -p litchi-numbers -p litchi -p litchi-cli`.
- Goldens blessed in both sets (`--all-features` and every feature but
  `encryption`): `--test capabilities --test registry`; `crud_matrix`,
  `corpus_fidelity`, `query_corpus_differential` blessed then re-run green;
  `python3 tools/officecli_capability_delta.py --officecli …/OfficeCLI`.
- `python3 tools/check_crate_boundaries.py` (67 packages, 282 internal
  dependency declarations); `python3 tools/non_iwork_gate.py verify` (46
  bulk roots, 42 safe facade trees, combined tree);
  `python3 -m unittest tools.test_non_iwork_gate` (59 tests).
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile
  ops-iwork` (new profile) and `--profile ops-odf`.
- `crates/pyo3-litchi/tests/run_tests.sh` (269 tests).
- `cd tools/perf-baseline && cargo check --locked --all-targets` (no
  lockfile change: the harness enables no iWork feature).
- The OfficeCLI corpus is unaffected (no iWork input in it).

## 8. Follow-ups

- **0297 Pages**: body tables, runs/formatting (request 13), paragraph
  add/remove (12), sections (11), section names, `/metadata` (14).
- **0298 Keynote**: notes writes, shapes (9), runs (10), slide add (8),
  tables; multi-slide fixture for `remove`/`move` evidence.
- **0299 Numbers**: formula writes (request 17), rows/columns (6),
  sheet/table lifecycle (4, 5), sheet/table move, cell styles (7).
- ADR 0034 W1 iWork renames (next wave).
- Remove the `ZipTail` extension point (`open_with_tail`, `_with_tail`
  detectors) once the facade's detection functions use the built-in iWork
  arm; it has no in-tree user besides the facade.
- Public prepared-source ingress (request 15) so `internal-iwork-source`
  stays facade-private; `Package::new()` (request 1).
- Share one object-tree core between `format::odf_tree` and
  `format::iwork` (the iWork copy keeps the two families independent).
- Batch coalescing of consecutive iWork sets on one leaf transaction.
