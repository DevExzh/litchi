# 0261 — XLSB W2 residue, source-backed metadata and P0 writes

- Status: implemented on `uo2/p7d-xlsb`, cut from `feat/unified-ops-api` at
  `949224c319` and merged with trunk (records 0258–0260, 0263). Plan
  [0248](0248-coordinator-restart-plan.md) §6 Phase B, step 5 (xlsb part)
  and step 7d.
- Date: 2026-10-13
- Scope: `crates/litchi-xlsb`, `crates/litchi-ops/src/format/xlsb*`, one
  additive engine change in `litchi-ooxml-common::properties`, consumer
  migrations (facade tests, perf-baseline harnesses), the glossary gate.
  `format/workbook` was not changed; the one shared ops hunk is a feature
  gate on `format/metadata.rs::round_trip` (trunk resolved it the same way).
- `performance_claim`: paired measurements in §4 (release, pinned CPU,
  ABBA, 3 trials). No frozen-harness qualification.

## 1. ADR 0034 W2 residue

**Ingress (`OpenOptions`).** One `litchi_xlsb::OpenOptions` envelope
(`source_limits`, `external_link_limits`, `leniency`, `parse_drawings`,
`cache_limits`, `execution_context`) replaces every positional combination:

- `Package::{open, from_bytes, from_reader, from_opc}_with_options`;
  `from_bytes` now borrows `&[u8]` and copies once into the owned source
  archive after the input-size check (so signed and exact-source publication
  is kept, as XLSX did in 0250); `from_slice*` is deleted.
- `Workbook::{new, from_opc_package}_with_options` (the projection keeps only
  these two configured constructors).
- `SourceBackedWorkbook::{from_reader, from_read_at}_with_options`.
- A policy an owner cannot honour is refused before input is read with
  `PackageError::UnsupportedOpenPolicy { policy: OpenPolicy }` (ops:
  `invalid_argument`), as PPTX did in 0252: cache and execution policies on
  owning routes, skipped drawing decoding on the source owner.
- Deleted without aliases: 14 `Package`, 9 `Workbook` and 14
  `SourceBackedWorkbook` ingress combinations
  (`*_with_limits`, `*_with_external_link_limits`, `*_with_leniency`,
  `*_with_cache_limits`, `*_with_execution_context`, `new_without_drawing_parse*`).

**Edit entries.** `Package::edit_with_options(structure::EditOptions)` and
`Package::edit_cell_values_with_options(selector, cell_values::EditOptions)`
replace `edit_with_limits(TransferLimits)` and the positional
`edit_cell_values_with_limits(selector, Limits, TransferLimits)`.

**Names.** `raw::Writer` is `raw::Encoder` (the `litchi-biff` precedent for a
record encoder; authoring stays `Builder`). The nested `ReadLimits` of
`data_model`, `calculation_chain`, `volatile_dependencies` and `xml_maps`
are `SemanticLimits`; the three `pub type Limits = ReadLimits` aliases are
deleted.

**`LimitKind`.** `cell_values::LimitKind::{SourceBytes, Cells, PatchBytes,
Changes, HistoryEntries, HistoryBytes}` names the dimensions of
`cell_values::{Limits, TransferLimits}`. Every caller-selected bound of the
cell-values, structure and worksheet families now reports
`PackageError::Limit { kind, observed, maximum }` (ops: `limit_exceeded`);
the cells/source-bytes bounds were `InvalidLength` (ops: `corrupt_file`)
before. `cell_values`, `cell_values::structure` and
`cell_values::worksheet` export all six family names (`worksheet` gains
`Diagnostics` and `Commit::diagnostics`) and are registered in both XLSB
glossary profiles.

**Gate.** `tools/api_glossary.json` registers 47 new bans per XLSB profile
(79 total) and three families (four with `metadata`). Glossary rows flipped:
transaction type, open from bytes, open from reader, configured open, open
envelope, ingress limit exceeded, limit kind enums, semantic limits; the
"xlsb cell_values pending" note is removed; a residue row lists what stays.

## 2. Source-backed metadata (record 0253 miss)

0253 measured a `/metadata` set on a source-backed XLSB handle at ×1.37–1.39
of the leaf, because the unified layer materialized the workbook first.

- **Engine** (`litchi-ooxml-common::properties`, additive):
  `Patch::compose(&next)`, `Patch::check_source_backed(&SourceBackedPackage)`
  and `Patch::prepare_source_backed(&SourceBackedPackage) -> PreparedTopology`.
  A patch becomes one source topology plan (part replacement, or part plus
  package-relationship addition/removal); stale sources are
  `Error::Conflict`, a changed patch of a signed source is `Error::Signed`,
  an empty patch is the exact source.
- **Leaf** (`litchi-xlsb`): `SourceBackedWorkbook::{metadata, app_props,
  edit_metadata, write_to, has_metadata_edits}` and
  `metadata::source::{Edit, Commit}`. A commit checks the composed patch
  against the source and the signature policy, and returns a new source
  owner that shares the catalog, every worksheet cache and every semantic
  read; the topology is planned once, on the first publication (`write_to`)
  or materialization (`to_workbook`). Sequential edits compose over the
  source.
- **Ops**: a `/metadata` set on an unsigned source-backed handle stays
  source-backed; publication streams through the leaf. Signed and
  removal-authorized handles keep the eager route.

## 3. P0 writes

**Cell formatting** (`set` on a cell, range or cell query). Setters:
`font.bold`, `font.italic`, `strike`, `underline` (`single`/`none`),
`font.name`, `font.size`, `font.color`, `fill` (solid; null removes),
`border.{all,left,right,top,bottom}` and their `.color`,
`numberformat`, `alignment.{horizontal,vertical,wrapText,textRotation,indent,
shrinkToFit,readingOrder}`, plus OfficeCLI aliases (`bold`, `bgcolor`,
`numfmt`, `wrap`, …). The leaf `cell_values::Format` is a partial change:
`Edit::set_format` copies the cell's `BrtXF` and its font, fill and border
records, patches only the named fields and interns the results (equal
records are reused), so the parent style, protection, font family/charset/
scheme and untouched resources keep their bytes. `AuthoredStyle` stays the
complete-style input for newly authored cells; both share the record
encoders. A blank coordinate becomes a formatted blank cell. Colors are
direct RGB; theme colors are `unsupported_capability`. The same properties
are now read (font keys only for a non-default font record, as OfficeCLI),
and their capability gaps are removed.

**Formulas.** `formula` text compiles through the leaf
(`Edit::set_formula_text`): references within the sheet, operators,
literals and built-in functions. The formula carries `fAlwaysCalc`, so Excel
recalculates on open; nothing is evaluated here. A value cell keeps its value
as the cache; a new cell caches an empty string. Formulas that name another
sheet or a defined name need the workbook's extern-sheet and name tables and
are `unsupported_capability`; invalid text is `invalid_value`. New formulas
are refused while an opaque calculation chain is present (as scalar edits
already were). Reads render binary operators parenthesized (`(A1*2)`), the
leaf's existing rendering.

**Merges.** `set(range, {merge: true|false})`.
`structure::Edit::{merge_cells, unmerge_cells}` (durable tag 11) rewrite
only the worksheet's `BrtMergeCell` block; a range overlapping a different
merge, a partial unmerge and single-cell ranges are refused. Ranges now
read `merge`. The reference commit `876588a23c` was used for its intent
only: its `fork_*_merged_cell_ranges` wrapped the mutable `Workbook` verbs,
which the transactional structure edit replaces.

**Defined names.** `add("/", namedrange, {name, ref, scope?})`,
`set("/namedrange[N]", {ref})`, `remove("/namedrange[N]")`, and reads of
`ref` and `scope`. `structure::Edit::{set_defined_name,
remove_defined_name}` (durable tag 12) splice `workbook.bin`: an existing
name keeps its flags and comment and receives a new `PtgRef3d`/`PtgArea3d`;
a new name is appended after the existing ones, so no `PtgName` index is
renumbered; the `Xti` is reused or appended, and a workbook without an
externals block gains one. Only the last name can be removed
(`unsupported_capability` otherwise). `named_ranges::Entry` /
`defined_name_entries` read names with scope and formula text.

**Behaviour per write** (tests in `crates/litchi-ops/tests/xlsb_p0_writes.rs`,
`crates/litchi-xlsb/tests/{cell_format_formula,merge_edit,defined_names_edit,
metadata_source}.rs`, `format/xlsb/metadata_tests.rs`): read-back on both
owners and after a byte round trip; exact no-ops keep the bytes; signed
workbooks refuse every write with `document_protected` until `unsign`; an
encrypted XLSB stays `Locked` (writes refused, unlock is
`unsupported_capability`) and re-encrypted output is refused without
publishing; transactions replay and invert byte for byte.

**Leaf fix.** Structure patches now publish their images through owned
sources, so applying (or inverting) a patch on an Excel-written workbook
reproduces its exact image (it re-serialized before).

**Capabilities.** `xlsb.json` capability and registry goldens re-blessed;
OfficeCLI delta: xlsb supported rows 62 → 78.

## 4. Measurements

Release builds, `taskset -c 17` on a shared 32-core host (load 8–13),
ABBA per pair, medians, three trials. Evidence (28 KiB):
`docs/unified-ops/results/0261/{meta,p0}-trial-{1,2,3}.jsonl`. Binaries are
not committed: `xlsb_metadata_source_bench` SHA-256 `eb98c900…ec26cad`
(tree of `5cbc45752d`), `xlsb_p0_write_bench` SHA-256 `7d9e061e…ceaf1ed`
(tree of `bc58fd78a8`); build:
`cargo build --release -p litchi-ops --features xlsb --example <name>`;
run: `taskset -c 17 target/release/examples/<name> 61` (metadata) / `31`.

**Metadata set** (`set("/metadata", {title})`; ratio = unified / leaf):

| Corpus | unified-source / leaf-source | unified-source / leaf-eager (0253's pair) | unified-eager / leaf-eager | set+publish, unified-source / leaf-eager |
|---|---|---|---|---|
| new workbook | 1.033–1.036 (19.6 µs) | 0.622–0.625 | 1.094–1.106 | 0.869–0.877 |
| Simple.xlsb | 1.021–1.028 (33.1 µs) | 0.616–0.619 | 1.105–1.109 | 0.859–0.872 |
| 8 × 2,000 cells | 1.018–1.025 (30.4 µs) | 0.230 | 1.056–1.060 | 0.514–0.521 |

The source-backed miss is closed: the default route is within 4% of its
leaf and 23–63% of 0253's eager leaf leg. The eager route is unchanged from
0253 (1.06–1.11, at the bar on small inputs; not addressed here).

**P0 single-op writes** (ratio = unified / leaf, three trials):

| Cells | Owner | value (ref.) | format | formula | merge | name |
|---|---|---|---|---|---|---|
| 10 | eager | 1.038–1.045 | 1.032–1.038 | 1.037–1.043 | 1.070–1.079 | 1.037–1.042 |
| 10 | source | 1.007–1.016 | 1.008–1.018 | 1.008–1.023 | 1.051–1.060 | 1.030–1.039 |
| 1,000 | eager | 1.005–1.009 | 1.007–1.010 | 1.005–1.019 | 1.039–1.053 | 1.019–1.021 |
| 1,000 | source | 1.004–1.008 | 1.007–1.010 | 1.005–1.014 | 1.030–1.039 | 1.015–1.020 |
| 10,000 | eager | 1.007–1.009 | 1.003–1.011 | 1.001–1.014 | 1.014–1.041 | 1.007–1.016 |
| 10,000 | source | 1.000–1.008 | 1.003–1.010 | 1.003–1.011 | 1.013–1.028 | 1.009–1.013 |

All within the 10% single-op budget. The source leaf leg opens the bytes as
a positional source (`from_read_at`), the route the unified byte open uses;
a source opened from a sequential reader keeps compressed passthrough and
its materialized edits are ~2× cheaper on small workbooks (161 µs against
360 µs at 10 cells) — a leaf follow-up (F4), not a unified overhead.

## 5. Gates (after the trunk merge, all exit 0)

| Gate | Command |
|---|---|
| format | `cargo fmt --all --check` |
| clippy | `cargo clippy -p litchi-ooxml-common -p litchi-xlsb -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` |
| clippy (xlsb subset) | `cargo clippy -p litchi-ops --no-default-features --features xlsb --all-targets -- -D warnings` |
| tests | `cargo test --all-features` for `litchi-xlsb`, `litchi-ooxml-common`, `litchi-ops`, `litchi` |
| rustdoc | `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-xlsb -p litchi-ooxml-common -p litchi-ops` |
| boundaries | `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` |
| glossary | `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (21 profiles, 1031 bans, 31 families) |
| delta | `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI> --check` |
| Python | `crates/pyo3-litchi/tests/run_tests.sh` |
| harnesses | `cargo check --all-targets --all-features` for every `tools/perf-baseline/unified-xlsb-*` |

`tools/perf-baseline` itself does not build on trunk (`litchi_ooxml_common::Props`
in `src/lib.rs`, removed by 0253); not touched here. Capability and registry
goldens for XLSB were re-blessed intentionally (§3).

## 6. Gaps and follow-ups

- **F1 owner model.** `Package::into_workbook` and the mutating `Workbook`
  verbs (`edit_opc`, `put_*`, `set_*`, `sign`) remain; scoped families keep
  `x_with_limits`/`edit_x_with_limits` read and edit entries and the crate
  error. `custom_data::Limits` (ooxml-common) is still re-exported.
- **F2 formulas.** Cross-sheet references and defined names in formula text
  (needs extern-sheet/name table authoring in the cell edit); formula
  removal (`formula: null`); `computedValue` (M4).
- **F3 formatting.** Theme colors, double/accounting underline (the leaf
  font model keeps a bool), superscript/subscript, fill patterns other than
  solid, diagonal borders, protection and `quotePrefix` remain gaps.
- **F4 source materialization.** `to_opc_package` of a positional source
  recompresses every part; keeping compressed passthrough would halve small
  source-owner edits.
- **F5 names.** Removing a name before others (needs `PtgName`
  renumbering), renaming, `comment`/`hidden`/`volatile` reads and writes.
- **F6 metadata.** `apply_metadata` on the source owner (durable replay of a
  patch onto a revised source) is not provided; the eager route ratio
  (1.06–1.11) is unchanged.
