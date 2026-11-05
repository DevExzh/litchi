# 0347 — Numbers swap, Python patches, the XLSX picture double read, the Patch docs and the RTF Mac code pages (second-audit package C)

- Status: implemented on `uo2/cross-small`, cut from `feat/unified-ops-api`
  at `926cbde41e`. Trunk was merged at `1f5a5442f4` (record 0344, tools
  and docs only; no crate changed). The delta, corpus and bench matrix were
  regenerated and checked after the merge.
- Date: 2026-10-19
- Input: the second definition-of-done audit of 2026-10-19 (§1 item 2a,
  §5, §6, §12.1 items 3, 8 and 9); records 0341 (G1, G2, G4) and 0338
  (follow-up 0338-P1).
- Owned:
  - `crates/litchi-ops/src/format/iwork/numbers.rs` and `crates/litchi-numbers`
    (one typed relocation refusal);
  - `crates/pyo3-litchi` with `python/litchi_py/__init__.pyi`;
  - the XLSX picture write path (`crates/litchi-xlsx/src/drawing/edit.rs`,
    `format/xlsx/feature_write.rs`);
  - the `transaction.rs` `Patch` docs;
  - `crates/litchi-rtf` (the code-page refusal).
- Small hunks outside them:
  - `capabilities.rs` and `capabilities/tests.rs` (the iWork swap arm);
  - `format/iwork/write.rs` (the Numbers swap dispatch arm);
  - `write.rs` (one doc paragraph on `File::swap`);
  - `error.rs` and `error/tests.rs` (the RTF code-page classification);
  - `tests/corpus_fidelity.rs` (the two RTF refusal rows);
  - `tools/unified_bench_matrix.py` (a `p0347` parser and the three swap
    scenario labels, so `--check` collects this record's evidence).
- `performance_claim`: measured (§6). Numbers swap is within 1% of the leaf
  moves. The XLSX picture `set src` falls from 1.27 to 1.07 of the leaf.

## 1. Numbers `swap`

Numbers was the only format that refused a definition-of-done verb for the
whole format (audit §6). Its golden named a follow-up,
`IWORK_API_REQUESTS.md`, that held no such request.

`swap` now exchanges:

| Operands | Leaf steps, each one committed transaction on the previous step's package |
|---|---|
| two `sheet`s | `edit_sheet_order`: the first sheet to the second's position, then the second (now one place earlier) to the first's |
| two `table`s of one sheet | the same two moves through `edit_table_order` |
| two `table`s of two sheets | `edit_table_relocation` of each into the other's sheet (the leaf appends), then one `edit_table_order` move in each sheet to the other table's old position |

- **Semantics.** As in XLSX and Keynote, a swap equals the two moves
  applied one at a time. `numbers_unified::swaps_sheets_and_tables` checks
  that byte for byte against the unified `move_to` calls: two moves for
  sheets and for tables of one sheet, two relocations and two moves across
  sheets. Swapping the operands gives the same bytes. An element swapped
  with itself (`/sheet[2]` and `/Sheet 2`) is an exact no-op.
- **Atomicity.** The steps are pushed onto one `Chain`. The unified commit
  publishes all of them or none, and its patch inverts to the exact source
  bytes in all three cases. In an atomic `edit`, a failing second swap
  leaves the snapshot byte-identical.
- **Same-named tables.** Table names are unique within a sheet. Two
  tables of one name in two sheets (`Table 1` in both, the default) would
  collide at the first relocation. The first table therefore passes
  through a temporary `edit_names` rename (a fresh `<name> N` unused in
  either sheet), which the last step reverts.
- **The leaf's refusal.** A relocation into a sheet that already has a
  table of that name failed the leaf's candidate reopen, which surfaced as
  `relocation::Error::Verification` and so as `internal_error`. Plain
  `move` had the same problem. `edit_table_relocation` now checks the
  destination before any rewrite and returns the new
  `Error::DuplicateTableName { path }`; the unified layer maps it to
  `invalid_value`. This covers a `move` into such a sheet and a swap whose
  table meets a third table of its name. Leaf test:
  `table_relocation::a_destination_name_collision_is_refused_typed`.
- **Rows and columns** are refused (`unsupported_capability`). The leaf's
  `table::topology` transaction inserts and removes rows and columns but
  does not reorder them. A reorder would have to carry storage, headers,
  UIDs, borders, merges and every formula coordinate that reads the
  table. So the brief's "where `table::topology` allows it" admits none.
  Two new operation gaps (`move row`, `swap row`) name follow-up F1.
  Cells keep their element-level refusal.
- **Capabilities.** The iWork arm admits `swap` wherever the format has
  movable elements, so the Pages/Keynote special case is gone. The
  `capabilities/tests.rs` mirror follows. `numbers.json` `verbs.swap`
  becomes supported.

**CRUD checklist.** Numbers swap cells moved from `gap` (follow-up
`IWORK_API_REQUESTS.md`) to `format-verb`. The ok cells are `sheet` and
`table`, the refused ones are `cell`, `row` and `col`, and `workbook` and
`mergeCell` stay unreached. The summary row now reads 0 gap, 4 move/swap
ok and 6 refused (it was 7 gap, 2 ok and 3 refused).

## 2. Python `Commit.patch`

Record 0341 G1: the Python `Commit` carried no patch. So the inverse and
conflict halves of the ADR 0008 patch row, and the stale-patch half of the
composition row, could not be reached from Python.

- `Commit.patch` returns the change's `Patch`. It is `None` only for
  `File.add_part` and `File.raw_set`, which are not transaction steps
  (0336 F2).
- `Patch.apply(file)`, `Patch.inverse()` and `Patch.is_noop()` use the ADR
  0034 §3 names of `litchi_ops::transaction::Patch`. `apply` releases the
  GIL and maps a refusal to the usual typed exception (`invalid_input`, a
  `CorruptFileError`, for another snapshot).
- A Python `Edit` runs one Rust transaction per staged run. `add` and the
  reads close a run. Its one `Patch` holds every run's Rust patch in
  order, and `inverse()` reverses that order. Each run's patch starts from
  the previous run's result, and a Rust `apply` returns the retained
  endpoint, so the chain replays exactly.
  `PatchTests::test_an_edit_patch_spans_every_run` covers an edit with
  three runs (`add`, a `get`, a `swap`).
- `batch`, `apply`, `merge`, `refresh` and `import_csv` pass their Rust
  commit's patch through unchanged.
- Stubs: `Commit.patch` and `class Patch`, plus `"Patch"` in `__all__`.
  `test_stub_sync` is green.

**Pytest mirror** (`test_crud_mirror.py::Adr0008MirrorTest`). The
patch-inverse row is now replayed in full:

- `patch.apply(source)` gives the commit's bytes;
- `inverse().apply(result)` gives the exact source bytes;
- the patch applied to its own result is refused with the recorded code;
- a byte-identical, separately opened copy refuses or accepts the patch,
  as the Rust row reports.

The composition row also requires each edit's patch, applied to the other
edit's result, to be refused with the recorded codes. Result: atomic
failure 21, patch inverse 21, composition 21 (`LITCHI_MIRROR_VERBOSE=1`).

## 3. XLSX picture `set src` (0338-P1)

Path resolution read the worksheet's drawing elements to find
`picture[N]`, and the leaf `replace_image` read them again to find the
picture's image part.

- The new leaf `drawing::edit::Edit::replace_element_image(sheet, index,
  &Element, …)` takes the element its caller already read.
  `replace_image` now reads the elements and delegates to it.
- The image relationship is still resolved in the staged drawing: the
  relationship must target the element's image part. An element that is
  not the picture at `index` is therefore refused, and nothing changes. The
  doctest checks both.
- The unified set passes the element its path resolution produced.

## 4. `Patch` docs (0341 G2)

The `transaction::Patch` rustdoc said that a separately opened file is a
foreign snapshot "even if its document text or bytes match". The ODT,
OTH and ODM handles and the ODF object-format handles (`same_snapshot` in
`format/odt.rs` and `format/odf_tree.rs`) compare bytes instead, so a
byte-identical reopened copy takes the patch. `crud-adr0008.tsv` reports
that for odt, oth, odm, odg, odc, odi, odb and odf.

The docs now say that most formats are owner-bound, name the formats that
check content instead, and say that a file whose bytes differ is refused in
every format. `Patch::apply` points at them. The behaviour is unchanged:
it is safe, because the patch is checked against the content it was made
from.

## 5. RTF Mac code pages (0341 G4)

`\ansicpg10001` (Mac Japanese) and `\ansicpg10007` (Mac Cyrillic) have no
exact codec in `litchi-codepage`. Decoding them as Shift JIS or KOI8-R
would corrupt text silently (the `litchi-rtf` corpus test says so), so
they stay refused. But they were refused as `MalformedDocument`, which the
unified layer reports as `corrupt_file`.

- The parser now returns the new `Error::UnsupportedCodePage(page)` for a
  positive page without a codec. A negative page is still malformed.
- `Classify` maps it to `unsupported_capability`.
- The `litchi-rtf` corpus test pins the variant for both files, and
  `error::tests` pins the code.
- In the fidelity report the two files move from `corrupt_file` to
  `unsupported_capability`. They remain "reader gap" rows, with follow-up
  F2.

## 6. Measurements

Release builds, `taskset -c 7`, ABBA. Outputs were compared byte for byte
outside the clock.

**Numbers swap.** `examples/numbers_unified_bench.rs` gains `sheet_swap`,
`table_swap` and `table_swap_sheets`. The leaf leg is the leaf
transactions the swap composes, hand-written; the unified leg is one
`File::swap`. Each scenario has 21 ABBA rounds (84 samples per leg), and
the existing `sheet_move` and `table_move` rows are rerun as controls.
Results: [run 1](results/0347/numbers-swap-run1.json) and
[run 2](results/0347/numbers-swap-run2.json).

| Scenario | Leaf (run 1) | Unified (run 1) | Ratio run 1 | Ratio run 2 | Identical |
|---|---:|---:|---:|---:|---|
| `sheet_swap` (three sheets) | 370.8 µs | 372.9 µs | 1.006 | 0.988 | yes |
| `table_swap` (three tables) | 534.9 µs | 535.4 µs | 1.001 | 0.996 | yes |
| `table_swap_sheets` (four steps) | 448.8 µs | 448.5 µs | 0.999 | 1.008 | yes |
| `sheet_move` (control) | 2,204.5 µs | 2,206.8 µs | 1.001 | 1.000 | yes |
| `table_move` (control) | 229.4 µs | 228.7 µs | 0.997 | 1.000 | yes |

**XLSX picture `set src`.** `examples/p0_residue_bench.rs` was rerun
unchanged; its leaf leg is still `replace_image`, which reads the elements
once. Results: [p0-residue-bench.txt](results/0347/p0-residue-bench.txt).
Before this change the record 0338 file gave 1.266 and 1.259.

| Scenario | Leaf | Unified | Ratio run 1 | Ratio run 2 |
|---|---:|---:|---:|---:|
| XLSX picture `set src` | 16.6 µs | 17.7 µs | **1.067** | **1.071** |
| XLSX run `set bold` (not touched; 0338 miss) | 1.82 µs | 2.45 µs | 1.346 | 1.367 |

The other rows of that bench are not on a changed path. The metadata and
cell rows stay within ±0.01 of their record 0338 ratios. The run set reads
1.346 and 1.367 here, against 1.324 and 1.312 in record 0338: its leaf leg
is 1.8 µs, so 0.02–0.05 is about 40–100 ns. This record changes nothing on
that path.

**Not measured.**

- Python `Commit.patch`. It adds one `Py::new` per commit (and for an
  `Edit`, one `Vec` push per run). `Patch.apply` clones `Arc`-backed
  endpoints. Neither is a per-operation hot path.
- The RTF and `Patch` doc changes do not touch a read or write path.

## 7. Gaps and follow-ups

- **F1 (litchi-numbers):** a row and column move in `table::topology`. It
  must carry storage, headers, UIDs, borders, merges and every formula
  coordinate that reads the table. With it, `move row` and `swap row`
  become supported. These are the two new operation gaps.
- **F2 (litchi-codepage):** exact Mac Japanese (10001) and Mac Cyrillic
  (10007) codecs. Until then, both RTF files are typed
  `unsupported_capability` refusals.
- The XLSX run `set bold` miss (1.32–1.37) is unchanged and outside this
  package (audit §5, "none named").
- Python raw-layer verbs still return no patch (0336 F2: `raw_set` and
  `add_part` are not `transaction::Edit` steps).

## 8. Gates

Every exit status was checked on its own (no piped `tail`). All of these
ran on the branch before the trunk merge. The merge changed no crate, so
only the tool checks were rerun after it, and they were green again.

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` (also after the merge) | 0 |
| `cargo clippy -p litchi-numbers -p litchi-xlsx -p litchi-rtf -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features numbers / pages / keynote / xlsx / rtf] --all-targets -- -D warnings` (six runs) | 0 each |
| `cargo test -p litchi-numbers --all-features` | 0 |
| `cargo test -p litchi-rtf --all-features` | 0 |
| `cargo test -p litchi-xlsx --all-features` | 0 |
| `cargo test -p litchi-ops --all-features --lib` (948) and `--doc` (215) | 0, 0 |
| `cargo test -p litchi-ops --all-features --test …`, 163 test targets in 9 chunks | 0 each |
| `cargo test -p litchi --all-features` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers -p litchi-xlsx -p litchi-rtf -p litchi-ops -p litchi-py` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, then the same without bless and with every feature except `encryption` | 0, 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix` / `--test corpus_fidelity` | 0, 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (298 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (40 profiles) | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0, 0 |
| `python3 tools/check_feature_matrix.py` (31 ok) | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_unified_bench_matrix` (after the merge) | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check` (after the merge) | 0, 0, 0 |

The OfficeCLI corpus has no Numbers, RTF code-page or Python lines, and no
XLSX picture result changed, so the corpus needed no re-bless.
