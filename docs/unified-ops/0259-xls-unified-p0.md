# 0259 — XLS unified P0 read and write

Status: landed on `uo2/p12-xls` (branched from `feat/unified-ops-api`
`e467bcf4e8`). This is step 12 of the coordinator plan (record 0248 §6,
Phase C): XLS reaches P0 read and write through the unified API, and the
ADR 0034 W4 residue for XLS is closed except where noted below.

## What changed

Every unified write is one native `Package::edit` operation that the
facade adopts. The facade has no second mutable workbook model.

- **Sheets.** `File::add("/", sheet)`, `remove`, `move_to`, `swap` and
  `set(name)` are supported on eager and positional owners. The native
  operations are `Edit::{insert_sheet, remove_sheet, reorder_sheets}` in
  `cell_values::tabs`. They remap all workbook-global tab references:
  `BoundSheet8`, `Window1` (active and first tab), `RRTabId`, the internal
  `SupBook` count, `ExternSheet` XTI ranges and `Lbl` local-name `itab`. They
  also fix the `INDEX` and `ExtSST` absolute offsets and deselect the old
  active `Window2`. Removal is refused for chart, dialog and macro tabs, for
  tabs that contain drawings, objects, pivot views or consolidation records,
  for tabs that own a sheet-scoped name, and for tabs that a formula or name
  still references. Tab edits are also refused on protected or shared
  workbooks, and adding or removing tabs is refused when a macro project is
  present. A move or insertion that would split a 3-D reference range is
  refused. Inserted worksheets come from a `Builder` template. Semantic patches record
  `sheet.reorder`, `sheet.insert`, `sheet.remove` and `sheet.restore`, so
  `Patch::inverse()` is exact.
- **Sheet-name identity.** The new `litchi_xls::sheet_name` module (`key`,
  `equivalent`, `is_valid`) is the single Unicode caseless key: NFD, default
  case fold, NFD (ADR 0034 §8). The reader's duplicate check, selector
  lookup, rename, insertion, comments, sheet visibility and chart sheet
  names all use it. Unified `add` and rename report `InvalidValue` and
  `DuplicateName`.
- **Formulas.** `set(cell, formula)` authors a BIFF8 `FORMULA` record with
  `Edit::set_formula`. The record has no cached result, keeps the cell's XF
  and replaces any scalar record. Ranges broadcast the formula in one native
  commit. Reads expose `formula` (rendered without the `=`, with the same
  rendering as the eager reader) and `cachedValue` on both routes. The
  positional route falls back to the eager tables only for 3-D formulas,
  which need workbook-global tables.
- **Rows and columns.** Writes cover `height` and `hidden` on `row[N]` and
  `width` and `hidden` on `col[X]` (`Edit::set_row_height`, `set_row_hidden`,
  `set_column_width`, `set_column_hidden`). Height accepts 0–409 pt, stored as
  twips. Width is in characters (256ths), up to 255. `ROW` records are
  patched in place or created with `fUnsynced`/`fDyZero`. `COLINFO` runs are
  split and merged. `add` inserts empty rows and columns at an explicit
  position through the existing structural engine, and `remove` deletes only
  empty rows and columns.
- **Cell formatting.** `cell_values::FormatChange` covers horizontal and
  vertical alignment, wrap, indent, shrink-to-fit, the four border styles and
  the fill pattern with foreground and background colours.
  `Edit::set_cell_format` rewrites those fields of the cell's XF, reuses an
  identical XF or authors a new one, and repoints the cell. Colours must be
  entries of the workbook's 56-colour palette.
  Reads expose alignment, border, fill, font, number format and protection,
  resolved from the cell's XF and the workbook's font, format and palette
  tables.
- **Named ranges.** Defined names are read as `namedrange` nodes with name,
  scope, formula and hidden state.
- **Read overhead.** The tab catalog is borrowed instead of cloned once per
  resolved tab (`Package::sheet_catalog`, `Workbook::sheet_catalog`).
  Positional source fences run once per facade call, after the read. Grids
  are sorted slices instead of B-tree maps.
- **Registry and capabilities.** `SETTABLE` covers cell, sheet, row and col.
  `ADDABLE` covers sheet, row and col. `REMOVABLE` covers cell, sheet, row and
  col. The XLS verbs add `add`, `move` and `swap`. The formula, cachedValue
  and namedrange gaps are removed. The `xls.json` capability golden was
  re-blessed and the OfficeCLI delta table regenerated.

### W4 residue

The scoped families in `sheet_visibility`, `ole_object`, `real_time_data`,
`revision_records`, `user_names`, `picture_compression`, `chart_area`,
`web_pub` and `external_link` now name their transaction `Edit`. These names
were deleted without aliases:

- `Transaction`, `UserNamesTransaction` and `UserNamesPackageTransaction`;
- `rollback` and `is_changed`;
- the duplicate `Patch::is_empty` (`is_noop` remains);
- `sheet_visibility::Snapshot::transaction`.

They are banned in the `xls-minimal` glossary profile, which now has 51 bans.
The `cell_values` and `sheet_visibility` selectors carry a checked
`litchi_core::Position` instead of a raw `usize`. `From<usize>` and
`From<Position>` are provided.

## Remaining gaps and follow-ups

- **Named-range writes.** The leaf has no `Lbl` editor, so `namedrange` is
  read-only and stays an operation gap.
- **Replacing an existing formula record.** This is refused with
  `UnsupportedCapability`. Its inverse needs the original record, including
  `SHRFMLA`/`ARRAY`/`TABLE` dependants. Formulas that use 3-D or external
  references cannot be authored, because the leaf parser does not resolve
  them, and are refused with `InvalidValue`.
- **Formatting.** Font, number-format and protection properties are
  read-only. Writing them needs new `FONT`/`FORMAT` records. A colour that is
  not in the palette is refused, and palette edits are not supported.
- **Rows and columns.** Inserting rows or columns on a sheet that has
  `COLINFO` records is refused (an existing structural-engine limit).
  Deleting a row or column that holds cells is refused. Outline levels are
  read-only.
- **Sheet removal and moves.** Removing chart tabs or tabs that contain
  drawings is refused (see above), as is a move that splits a 3-D reference
  range.
- **Python.** No Python surface is changed by this record.
- `property_set::Transaction` remains, because the record 0254 metadata
  owner (`package/metadata.rs`, outside this scope) uses it.
- **Leaf findings, not fixed here:**
  - `Edit::insert_cell` gives new cells XF 0 rather than the default cell
    XF 15.
  - Inserted cell records are appended unsorted within their row.
  - The chart editor's `rewrite_sheet_directory` does not adjust `INDEX`
    offsets.

## Read overhead (record 0203 miss)

Record 0203 measured unified XLS reads at 13–36× native. Most of that gap
came from the benchmark, not from the adapter. The old native leg returned raw
`CellValue`s, while the unified leg formatted display text, parsed paths in
the timed loop, and ran on a different owner route. The rewritten
`xls_read_bench` (schema `litchi.xls_read_bench.v2`) gives both legs the same
output work (cell display text), uses pre-parsed paths, and measures eager
and positional owners in ABBA order. The adapter changes above then removed
the per-tab catalog clones and the repeated source probes.

Nine release samples on CPU 30 (AMD EPYC 9R45), median ns per complete scan:

| Cells | Owner | Native get | Unified get | Ratio | Native scan | Unified query | Ratio |
|---:|---|---:|---:|---:|---:|---:|---:|
| 1 | eager | 73 | 239 | 3.27× | 59 | 351 | 5.95× |
| 100 | eager | 5,908 | 26,150 | 4.43× | 4,959 | 14,272 | 2.88× |
| 1,000 | eager | 68,071 | 301,894 | 4.43× | 53,224 | 146,814 | 2.76× |
| 10,000 | eager | 828,584 | 2,885,573 | 3.48× | 584,352 | 1,509,267 | 2.58× |
| 1 | positional | 1,244 | 2,009 | 1.61× | 1,151 | 2,570 | 2.23× |
| 100 | positional | 109,344 | 203,164 | 1.86× | 11,772 | 73,650 | 6.26× |
| 1,000 | positional | 1,130,723 | 2,074,035 | 1.83× | 101,877 | 716,028 | 7.03× |
| 10,000 | positional | 15,566,431 | 25,027,173 | 1.61× | 1,027,505 | 7,270,283 | 7.08× |

On the same benchmark before the adapter changes (five samples, commit
`6d199c3f6c`), the ratios were: eager get 4.1–5.4×, positional get
2.5–3.1×, and positional query 10.7–12.0×.

**This is still a miss** for an overhead target near 1×. The remaining cost
sits in shared code outside this scope:

- **Eager get** costs about 200 ns per cell. The generic resolver moves the
  264-byte `Node` through several frames; memmove is about 23% of a profile.
- **Eager scan.** It still builds a node per match.
- **Positional query.** The generic walker (`query/walk.rs`) and resolver
  (`path/resolve.rs`) each fence every step. On a positional owner, every
  fence calls `FileSource::version()`, which takes a mutex and makes an
  `fstat` system call.
- **Early termination** (record 0203). It is still open, because a worksheet
  query materializes the sheet's sorted grid before it yields.

Follow-ups for the shared-code owners: a single fence per walk, smaller or
boxed node payloads, and a streaming cell cursor.

## Write overhead

`xls_p0_write_bench` (schema `litchi.xls_p0_write_bench.v1`) compares, on the
same input, native `Package::edit` plus one operation plus `commit` with the
single unified `File` call, which includes target parsing and property
coercion. Serialization is not timed. The input is a three-sheet workbook
with N numbers on `Data`. Both legs' outputs were serialized and compared
first, and every scenario published **byte-identical** output. Nine samples,
ABBA order, CPU 30, ratio of medians (unified / native):

| Scenario | eager 100 | positional 100 | eager 10k | positional 10k |
|---|---:|---:|---:|---:|
| sheet_move | 1.009 | 1.026 | 1.004 | 1.000 |
| sheet_add | 1.008 | 1.038 | 1.000 | 1.001 |
| formula_set | 1.024 | 1.064 | 1.005 | 1.001 |
| row_height | 1.029 | 1.120 | 1.020 | 1.099 |
| column_width | 1.033 | 1.126 | 1.021 | 1.105 |
| cell_format | 1.025 | 1.072 | 1.002 | 1.003 |

Native medians at 10k cells range from 3.9 ms (positional add) to 8.8 ms
(eager row height). Row and column writes on a positional owner miss by
about 10%. Resolving the `row[N]`/`col[X]` target through the generic
resolver projects the sheet's axis layout from the source. This costs about
600 µs at 10k cells and is part of the same resolver follow-up.

## Evidence

The reduced JSON results are under
[`results/0259-xls-unified-p0/`](results/0259-xls-unified-p0/), each under
20 KiB:

- `read-before.json`: five samples, built at `6d199c3f6c`;
- `read-after.json`: nine samples;
- `write-p0.json`: nine samples.

No binary is committed. Measured binaries, built from this record's commit
with `cargo build --release -p litchi-ops --features xls --example <name>`
and run as `taskset -c 30 target/release/examples/<name>`:

- `xls_read_bench`:
  `200bc59a02c12076338071db85b873c026c8a9f1dc6a049da2f98a27f04611ec`
- `xls_p0_write_bench`:
  `1a7ffe6069871eb914bfaf32b52ba5fcab42fff85fd1fd1edad561a6cc782afd`

## Tests

The new integration tests in `crates/litchi-ops/tests/` are
`xls_unified_structure.rs`, `xls_unified_formula.rs`,
`xls_unified_layout.rs`, `xls_unified_format.rs` and
`xls_unified_protection.rs`. `xls_unified_read.rs` and
`xls_unified_remove.rs` are extended. Together they cover read-back on both
owner routes, exact no-op, exact inverse through `Patch::inverse`, batch
attribution, typed and atomic refusals, and signed and unlocked-encrypted
refusals (`DocumentProtected`) for every P0 write. Leaf unit tests are in
`cell_values/{tab,formula,layout,format}_tests.rs`. They include a corpus
round trip that reorders tabs and checks 3-D formula values.
