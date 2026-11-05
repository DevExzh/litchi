# 0370i — W1-A slice 10: Numbers builder table dimensions, language and locale

- Status: landed on branch `uo2/iwa-w1a`, after slice 9
  ([0370h](0370h-iwa-w1a-table-duplicate.md), merged on trunk).
- Date: 2026-10-26.
- `performance_claim`: the leaf builder with options runs within the 5% budget, at +2.4% Ir over the host builder (§4).

## 1. What landed

`litchi_numbers::Builder` gains three options:

| Option | What it sets | Default |
|---|---|---|
| `table_dimensions(rows, columns)` | the sheet table's size, headers included | 10 × 5 |
| `language(tag)` | a BCP 47 tag | `en` |
| `locale(id)` | the locale | `en_US` |

There are two new error variants: `create::Error::InvalidLanguage` and
`InvalidLocale`. A size the table cannot take fails with
`create::Error::Dimensions`.

**Where the values go** (the host builder writes the same fields):

- Names: the sheet's `name` and the table model's `table_name`.
- Dimensions: the model's size; the tiles' extent (`maxColumn`,
  `maxRow`) and their removed rows; the header buckets' headers inside
  the new size; a fresh row/column UID map (the template table stores no
  cell); the stroke sidecar's size; and the table owner's table and body
  extents.
- Language: `TSA.DocumentArchive.document_language`.
- Locale: the document's `locale_identifier` and
  `creation_locale_identifier`, and the calculation engine's
  `saved_locale_identifier`.

**How the build writes them.** `package/builder_synthesis.rs` writes every
option, with the identity, straight into the template in one pass:

- It scans the template's two changed components (`Index/Document.iwa`
  and `Index/CalculationEngine.iwa`) object by object, reading only the
  headers' identifiers, message types and lengths.
- It replaces the affected payloads and their header lengths.
- It reassembles the package once and opens it once. The final open
  validates the result in full.

No structural transaction publishes or reopens in between. The package
metadata is untouched, so the output is one first save with the
template's save tokens. `package/document_locale.rs` holds the language
and locale wire rewrites.

A build without options returns the template, or the template under the
requested identity, as before.

**Host changes:**

- *Defect fixed* (mac-queue row): the host builder's table owner carried no
  table or body extents. Native owners record both, and so have tables
  added by the editor since record 0370f. The Pages builder borrows this
  graph for its first table and keeps the extent-less owner
  (`Builder::without_owner_extents`): the Pages body-table hidden-axes
  route refuses an owner with spanning extents. That is W1-B's to lift.
- *Alignment, not a defect:* the host builder's build-history producer
  string is now the focused crate's `Created by Litchi` (record 0309).

**Documented difference:** a table with no body row or column (1 × 1).
The host builder writes one; the leaf refuses it, as `table::resize`
does, because a native table keeps at least one body row and column.
`numbers_builder_headers_only_refusal` pins this.

Nine ledger rows are migrated:

- `table_dimensions`, `language` and `locale` → `litchi_numbers::Builder`;
- `build`, `builder`, `new`, `sheet_name` and `table_name` →
  `litchi_numbers::Builder`;
- `create` → `litchi_numbers::Package::new`.

## 2. Parity

`crates/litchi-iwa/tests/parity/numbers_builder_cases.rs` builds with the
host builder and then with the leaf builder under the host output's
identity. That identity is the document, version and private UUIDs read
from its properties.

Both tests are tier B. The host draws these values at random, so they are
compared by form, as a documented policy:

- the table UUID, the owner UUID and the owner-map entry derived from it;
- the row/column UIDs (lengths and permutations);
- the registry's object UUIDs.

| Test | Cases |
|---|---|
| `numbers_builder_table_dimensions_parity` | 10×5, 20×8, 3×2, 2×2, 40×26 |
| `numbers_builder_language_locale_parity` | fr/fr_FR, ja/ja_JP, de-CH/de_CH |
| `numbers_builder_names_parity` | default; sheet only; table only; both, with non-ASCII names |

## 3. Gaps

| Item | Next step |
|---|---|
| no native evidence for builder options | mac queue |

## 4. Performance

`cost_probes::numbers_builder_cost_abba` builds a 20 × 8 `fr`/`fr_FR`
workbook in the order host, leaf, leaf, host:

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| builder with options | 2,500,108 | 2,561,096 | **+2.4%** |

The first version composed the names, resize and locale transactions,
each publishing and reopening the package. It cost 9,304,911 Ir, +275%
over the host. The one-pass synthesis removed those intermediate
publications. It also removed two quadratic encoders:

- the shared `litchi_numbers_wire::table_create::empty_storage` UID-map
  encoder re-parsed its growing output on every append; it now encodes in
  one linear pass, with the same bytes, which also speeds up the
  table-add route;
- `table_topology::geometry` reread the model once per field; the
  synthesis reads the geometry in one pass.

What remains is mostly the one final open (1.42 M Ir), which every leaf
route that returns a `Package` pays.

## 5. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` | 0 |
| `cargo test --all-features -p litchi-numbers`, `-p litchi-numbers-wire` | 0 / 0 |
| `cargo test --all-features -p litchi-pages`, `-p litchi-keynote` | 0 / 0 |
| `cargo test -p litchi-ops --all-features --test numbers_unified` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
