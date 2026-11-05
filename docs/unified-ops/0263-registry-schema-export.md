# 0263 — Registry schema export and golden snapshots

Date: 2026-10-13. Status: implements record 0248 §6 Phase D step 16.
Branch `uo2/p16-schema`, cut from trunk `fb972fa9ed`.

The property registry is now exported as data, golden-tested per format, and
the OfficeCLI delta table is generated from that export. The delta now
covers every format the registry covers, so DOC, PPT and RTF join it
(record 0249 follow-up).

## 1. The export

`litchi_ops::schema(format) -> schema::Schema` is re-exported by the facade
as `litchi::schema`. It wraps `Capabilities::of(format)` and reuses its
`Element`, `Property` and `Gap` types, so there is no second element model.
`Schema::{format, registered, elements, element, gaps}` are allocation-free;
`serde::Serialize` writes JSON.

The JSON uses OfficeCLI's help-schema shape for each element (`element`,
`parent`, `elementAliases`, `operations`, `paths`, `children`, `properties`).
Each property carries:

| Member | Source |
|---|---|
| key | the OfficeCLI key the node view prints and `set` accepts |
| `type` | `ValueType::schema_type` (`bool`, `int`, `number`, `string`, `scalar`, `enum`, `length`, `font-size`, `color`) |
| `values`, `valueAliases` | the `Tokens` vocabulary: canonical values and accepted input spellings |
| `unit` | lengths: the unit a bare number is read in (`emu`, `pt`, `twip`), from the `LengthContext` |
| `colorContext` | colors: `drawing`, `spreadsheet` or `word` |
| `aliases` | the registry table's aliases (`font.strike` → `strike`) |
| `constraint` | `{min,max}` (int, number), `{minEmu,maxEmu}` (length), `{maxChars}`, `{notEmpty}` |
| `clearable` | `null` clears the property |
| `add`, `set`, `get` | the access flags, as in `capabilities()` |
| `writeGap` | `{code, followUp, reason}` for an `Access::Gap` registry entry |

Each element also carries `propertyGaps` (keys OfficeCLI has that Litchi
does not answer yet, with their follow-ups), `predicates` and its own `gap`.
The top-level `gaps` list holds element gaps that have no table entry, plus
axis, operation and view gaps. A format whose adapter is not in the build
has no elements and a `pending` entry instead.

Plumbing:
- `PropSpec` gains one field, `admission: registry::Admission` (constraint,
  clear flag, pending-setter gap). `Entry::readable` fills it, so every
  projection a write table produces carries its declaration.
- `registry::Constraint` becomes public as `capabilities::Constraint`
  (`#[non_exhaustive]`). Its equality compares `f64` bounds by bit pattern.
- `Table::prop_aliases()` is a `const` projection of a table's aliases.
  `schema::prop_aliases` maps (adapter, element) to it. Only the XLSX cell,
  range, row and column tables declare aliases today. The test
  `aliased_tables_are_exported` lists every aliased table; a new one adds its
  arm there.
- `capabilities::Property` gains `aliases()`, `constraint()`, `clearable()`
  and `write_gap()`.

## 2. Golden snapshots

`crates/litchi-ops/tests/registry.rs` compares
`serde_json::to_string_pretty(schema(format))` with
`tests/golden/registry/<format>.json` for doc, docx, ppt, pptx, rtf, xls, xlsb
and xlsx. `LITCHI_BLESS=1` regenerates them. This mirrors the capability
goldens: with the `encryption` feature, DOCX and XLSX are written as
`<format>-encryption.json`, because that feature adds operation gaps. So a
build without `encryption` is what checks the plain DOCX and XLSX files.
Ten files were blessed, 14–81 KB each (532 KB in total).

A second test, `registry_export_matches_capabilities`, checks two things for
every format:
- the export and `Capabilities::of` agree on every element and on every
  property's flags;
- no alias equals a property name of its element.

## 3. Defects found and fixed in per-format tables

1. **XLSX cell capabilities omitted four writable keys.** `border`,
   `border.all`, `border.color` and `border.all.color` are in
   `CELL_WRITE_TABLE`, and cell `set` accepts them
   (`tests/xlsx_cell_writes.rs`, `border_shorthands_and_style_aliases_*`).
   But `cell_capability_array` appended only `value` and `clear`.
   - It now appends every write entry that is not a readable cell property,
     write-only (`get: false`).
   - The capability goldens `xlsx.json` and `xlsx-encryption.json` were
     re-blessed. That is the only capability golden change: +4 cell
     properties each.
   - In the delta table, XLSX property writes move from 11/36/497/66 to
     11/39/494/67 (supported/partial/gap/litchi-only), and XLSX property reads
     gain one litchi-only row.
2. **Flag merges dropped the registry declaration.** The xlsx, xls and xlsb
   capability projections copy `access.set`/`access.add` from a write entry
   onto the read property of the same name. They now also copy `admission`,
   one line per merge site: 6 in xlsx, 2 in xls, 2 in xlsb. Without it, a
   merged property (for example XLSX `font.charset`, `row.height`) would
   have exported no constraint and no clear flag.

Shared-file hunks (other agents edit these files in parallel):
- `PropSpec` struct literals gain `admission: Admission::NONE`: rtf (23
  literals), pptx (1 macro arm), ppt (1), doc (1).
- `capabilities.rs`: the `Constraint` re-export, the `Data.prop_aliases`
  member and its fill, the four `Property` accessors, `header_columns`
  becoming `pub(crate)`, and one module-doc paragraph.
- `format/xlsx.rs`: a separate `pub(crate) use` line for the three aliased
  tables.

Merge hazard: any `PropSpec { .. }` literal added on another branch must also
gain `admission: crate::registry::Admission::NONE`, or the merge will not
compile. `PropSpec::new` callers are unaffected.

## 4. Delta table

`tools/officecli_capability_delta.py` now reads
`crates/litchi-ops/tests/golden/registry/*.json`:

- The formats are every snapshot present, in the order docx, doc, rtf, xlsx,
  xlsb, xls, pptx, ppt, then any other format sorted.
- DOC and RTF are compared against DOCX's help vocabulary, XLS and XLSB
  against XLSX's, and PPT against PPTX's (`REFERENCE`).
- Property gaps come from each element's `propertyGaps`.
- A format is pending when the export has `pending`.
- An optional `<format>-encryption.json` contributes operation gaps for any
  format; before, this was DOCX only, so XLSX gains its encryption operation
  gaps.
- An OfficeCLI element with no Litchi table entry but a top-level element gap
  (the 22 DOC element gaps of record 0249) is now `gap` with the follow-up,
  and its properties are `element gap`, not `missing`.

`--check` and the unit tests are kept. There are now 24 tests (19 before):
- new: family vocabularies, snapshot discovery, undeclared element gaps;
- reworked: the golden-hash test now covers every snapshot.

The regenerated table has 7,453 lines and 558,791 bytes, under the 1 MiB
guard. `--check` exits 0 against both the vendored `test-data/officecli` and
`/home/zhuhe/code/litchi/3rdparty/OfficeCLI` (commit `dced0d74`; both trees
hash the same).

DOCX, XLSB, XLS and PPTX rows are unchanged by the switch of input, which
shows the registry export and the capability goldens agree. XLSX changes
in two ways only:
- defect 1 (3 cell rows change, 1 is added);
- its encryption operation gaps, now merged: 4 contextual rows.

New rows (OfficeCLI properties with add/set; supported/partial/gap):

| Format | Elements supported/gap/missing | Property reads supported | Property writes supported/partial/gap |
|---|---|---|---|
| doc | 15/11/18 | 39 | 1/13/782 |
| rtf | 7/12/25 | 25 | 12/9/775 |
| ppt | 13/20/0 | 31 | 0/12/549 |

## 5. CLI and Python

- **CLI:** `litchi schema <type> [--json]`.
  - `--json` prints the serialized `Schema`.
  - Text output lists each element with each property's type, `asg` flags,
    aliases and write gap.
  - `tests/cli.rs::schema_json_is_the_registry_export` checks equality with
    the in-process export for five formats, checks the text output, and
    checks exit statuses 2 (no argument) and 4 (`txt`).
- **Python:** `litchi_py.schema(format)` accepts a `FileFormat` or a name
  and returns the dict through `pyser::to_py` (no JSON text). The `.pyi`
  gains `def schema` and the `_Schema` TypedDict.
  - `SchemaTests` checks that the dict equals the Rust registry golden for
    all eight formats, member order included.
  - It checks the TypedDict keys and that `PropertyError` is raised for
    unknown names.
  - It does not release the GIL: the call takes tens of microseconds (§6).

## 6. Measurements

The export is not a hot path; this is its order of magnitude.

Benchmark setup:
- Example: `crates/litchi-ops/examples/schema_export_bench.rs`, run with
  `cargo run --release -p litchi-ops --all-features --example
  schema_export_bench` under `taskset -c 3`.
- Machine: AMD EPYC 9R45, kernel 7.0.0-1013-aws.
- Each figure is the median of 31 samples; two runs agreed within 1.5%.

| Format | Properties | JSON bytes (compact) | `schema()` | typed walk | `serde_json::to_string` |
|---|---|---|---|---|---|
| doc | 44 | 14,417 | 1.5 ns | 1.84 µs | 19.8 µs |
| docx | 185 | 43,178 | 1.5 ns | 8.20 µs | 63.3 µs |
| ppt | 39 | 23,093 | 1.5 ns | 2.10 µs | 44.5 µs |
| pptx | 232 | 44,416 | 1.5 ns | 11.11 µs | 75.7 µs |
| rtf | 23 | 10,154 | 1.5 ns | 0.81 µs | 10.2 µs |
| xls | 18 | 20,139 | 1.5 ns | 0.64 µs | 31.0 µs |
| xlsb | 65 | 35,655 | 1.5 ns | 2.24 µs | 50.8 µs |
| xlsx | 163 | 42,000 | 1.5 ns | 8.43 µs | 59.2 µs |

`PropSpec` grows by `Admission`, and the read paths copy and iterate
`PropSpec`. So the existing paired read probes were rerun on a baseline build
(this branch with its tracked changes stashed) against the new build.

Read-probe setup:
- Probes: `ppt_read_bench` and `xls_read_bench`, built with
  `--no-default-features --features ppt,xls`, release, `taskset -c 3`.
- Order: base, new, new, base, base, new, new, base. For xls,
  `LITCHI_XLS_BENCH_SAMPLES=9`.

Medians of all samples:
- Unified legs:
  - `ppt` `unified_get`: 0.972–1.003 of baseline.
  - `ppt` `unified_first`: 0.990–1.022.
  - `xls` `unified_get`: 0.952–0.966 at 100–10,000 cells; 400 → 430 ns at
    one cell, where the timer step is 10 ns.
  - `xls` `unified_query`: 0.988–1.009.
- Native legs moved 0.961–1.016 in the same runs (they do not touch
  `PropSpec`), which bounds the noise.

No regression above 5% beyond that one-cell quantization. The raw output is
not committed (scratch only).

## 7. Verification

Environment: toolchain 1.95.0, `CARGO_TARGET_DIR=<worktree>/target`,
`CARGO_INCREMENTAL=0`, debug info off.

| Command | Exit | Result |
|---|---|---|
| `cargo fmt --all --check` | 0 | |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 | |
| `cargo clippy -p litchi-ops --lib --tests [--features docx\|rtf] -- -D warnings` | 0 | the featureless build and docx-only and rtf-only builds |
| `cargo test -p litchi-ops --all-features` | 0 | 1,809 passed |
| `cargo test -p litchi-ops --features doc,docx,ppt,pptx,rtf,xls,xlsb,xlsx --test registry --test capabilities` | 0 | the plain DOCX/XLSX goldens |
| `cargo test -p litchi --all-features` | 0 | 260 passed |
| `cargo test -p litchi-cli --all-features` | 0 | 18 passed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops -p litchi -p litchi-cli -p litchi-py` | 0 | |
| `python3 tools/check_crate_boundaries.py` | 0 | |
| `python3 tools/non_iwork_gate.py verify` | 0 | |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir <wt>/target/glossary` | 0 | 21 profiles, 900 banned names, 24 families |
| `python3 -m unittest tools.test_officecli_capability_delta` | 0 | 24 tests |
| `python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check` | 0 | |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 | 245 tests |

Pre-existing, not from this change: `cargo clippy -p litchi-ops --lib
--features xlsx -- -D warnings` fails on dead code in an XLSX-only build
(`round_trip`, `crates/litchi-ops/src/format/metadata.rs:362`). This branch
does not touch that file.

## 8. Not claimed; follow-ups

- **Aliases come only from `Table` alias arrays.** If a writer accepts a
  spelling outside its registry table, the export does not list it. Today
  only XLSX tables declare aliases. The DOCX, PPTX, PPT and RTF registries
  declare none, though OfficeCLI lists aliases such as `strikethrough`
  (owner: each format's registry wave).
- **The constraint is shown on the read type.** On a merged property the
  constraint and type can come from different declarations. XLSX `fill`
  exports `type: color` (read) with `constraint: {notEmpty}` from its
  text-typed write entry. A follow-up could export the write type
  separately.
- **No prose members.** OfficeCLI's `description`, `examples`, `readback`
  and `enforcement` have no registry source and are not exported.
- **No meta-schema.** No JSON Schema file describes Litchi's additional
  members; the module documentation is the contract.
- **The plain DOCX and XLSX snapshots need a run without `encryption`.**
  The `--all-features` gate checks only `*-encryption.json`. The same holds
  for the capability goldens.
- **Python `schema()` keeps the GIL** (about 10–80 µs).
- **ODF and iWork are not exported yet.** They have no registry; they
  appear as `pending` exports and join the delta when their adapters
  declare tables (0248 Phase F).
