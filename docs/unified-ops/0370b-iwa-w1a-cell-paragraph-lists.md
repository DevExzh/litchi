# 0370b — W1-A slice 3: Numbers cell paragraph lists

- Status: landed on branch `uo2/iwa-w1a`, after slice 2
  ([0370a](0370a-iwa-w1a-shared-cell-writer-and-cell-text.md), merged as
  `22ea14371f`); trunk merged afterwards.
- Date: 2026-10-26.
- `performance_claim`: the new cell-list route runs 84.0% below its host
  route (§4).

## 1. What landed

| Commit | Crates | What |
|---|---|---|
| `96af8a4d53`, follow-up | `litchi-numbers-wire`, `litchi-numbers` | shared cell writer: a Numbers row after every stored row follows the last row record; metadata transitions that only add registry records, or add and withdraw the same edge (§3) |
| `80d39a889d` | `litchi-numbers`, `litchi-iwa`, tools, docs | **new family `litchi_numbers::cell::list`**, the storage-list publisher `package::storage_list`, two host defect fixes; 31 ledger rows migrated (§2) |
| this commit | docs | this record |

## 2. `litchi_numbers::cell::list`

The API is `Package::{edit_table_cell_lists, apply_table_cell_lists}` with
`Edit::{set_list, set_lists, set_level, set_numbering, set, reset}`, and
the reads `Package::{table_cell_list, table_cell_lists,
table_cell_list_levels, table_cell_list_level, table_cell_list_numbering,
table_cell_list_property}`. The host's 31 methods collapse into these
verbs: `reset_table_cell_paragraph_list` is `set_list(None)`,
`reset_table_cell_paragraph_list_level` is `set_level(ZERO)`, and the seven
`set_/reset_table_cell_paragraph_list_<property>` pairs are `set` and
`reset` with W1-D's `list::Property` / `PropertyKind`. The value types are
re-exported from `litchi-iwa-text` (an allowed public dependency), which
`litchi-numbers` now depends on (`tools/crate_boundaries.json`).

**Promotion.** A plain cell (empty, or a string key) reads as unlisted
paragraphs. The first change that needs a list promotes it as the host's
`cell_paragraph_list::promote_plain_cell` does: the table gains a
rich-text list when it has none, a `TSWP.StorageArchive` (kind CELL, the
cell's text, its text style and the stylesheet's `None` list preset at
zero, one paragraph-data, start and bidi entry per paragraph) and a
`TST.RichTextPayloadArchive` naming it are stored, a rich-text entry names
the payload, the cell becomes a rich-text cell and its string entry is
released. A change that leaves a plain cell unlisted (`set_list(None)`,
level zero, continued numbering, a reset) is a no-op after a paragraph-start
check, as on the host. A rich-text payload shared by several cells is
refused (`UnsupportedDependency`); the host copies it first.

**Publication.** `package::storage_list::Publisher` publishes W1-D's
`litchi_iwa_text_wire::list` plans on any storage of the working set, as
the host text editor does: presets found or stored beside the stylesheet,
paragraph variations stored and registered under their parent, list-style
boundaries rewritten with their reference sequence, and the registry's
UUIDs and cross-component references kept in step (a reference no object
of the storage's component uses any more goes). W2-A can run it on Pages
and Keynote storages.

**Host defects fixed** (mac-queue row):

1. Promotion failed on every native table: native Numbers keeps a table's
   rich-text list in its own `Index/Tables/DataList-*.iwa` component, and
   the host updated the table model's component, so no fixture could be
   promoted. The payload and its storage now go beside the list (a new
   list goes beside the table model).
2. Promotion registered no external reference from the list's component to
   the storage's stylesheet, paragraph style and list preset, which live in
   `DocumentStylesheet.iwa`.

**Parity** (`crates/litchi-iwa/tests/parity/numbers_cell_list_cases.rs`), on
up to three cells of the first table of `basic.numbers`,
`table-data-list-native.numbers` and `custom-text-native-unicode.numbers`
(empty and text cells):

| Test | Routes | Tier |
|---|---|---|
| `numbers_cell_list_preset_parity` | 90: bullet and numbered presets and one-placement lists on plain cells, then reset, `None` and the other preset on the host's output | B |
| `numbers_cell_list_level_numbering_parity` | level on plain and listed cells, numbering, level reset | A |
| `numbers_cell_list_property_set_parity` / `_rewrite_` / `_reset_` | the seven properties set (new variation), rewritten in place, reset (collapse) | B / A / B |
| `numbers_cell_list_paragraph_parity` | a two-paragraph cell written by the leaf's cell writer: mixed placements, second-paragraph level, numbering and properties | A and B |
| `numbers_cell_list_refusal_parity` | a list property of a plain cell; a level at a non-start; one preset over mixed paragraphs | both refuse |
| `numbers_cell_list_read_parity` | every list read on corpus cells and on every listed cell above | A |

Tier B covers the fresh UUID of a new list style and the leaf's save-token
policy. The ledger migrates 31 rows; the glossary registers the family in
`numbers-minimal` and `numbers-full` and in `docs/API_GLOSSARY.md`.

## 3. Shared engine changes (`96af8a4d53`)

- `table_write::tile_set_cell_in_field_order` (Numbers): a new row after
  every stored row is written after the last row record and before the
  tile's storage-version fields, as the Numbers host writer's encoder
  places it. `tile_set_cell` keeps appending such a row after them, as the
  Pages and Keynote host writers do (`pages_body_table_cell_parity` and
  `pages_body_table_cells_batch_parity` pin that); the Numbers
  materialisation of an empty cell (`cell::text_style`, `cell::text`,
  `cell::list`) uses the field-order form.
- `structure::metadata::Transition::remove_reference` also withdraws an
  edge the same transition added (a preset referenced, then left unused).
- The package-metadata addition batches require a rising watermark. A
  transaction that adds registry records without allocating an identifier
  (a reference to an existing preset) now passes the batch one identifier up
  and restores the watermark afterwards.
- A stray backtrace `eprintln!` in `table_topology::shift` (record 0329) is
  gone.

## 4. Performance

Callgrind on the release parity binary, CPU 3 pinned,
`cost_probes::numbers_cell_list_cost_abba` (host, leaf, leaf, host on each
input).

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| cell list: bullet on plain B2 of `basic.numbers` (promotion), numbered on the host's output | 1,715,520,591 | 274,527,749 | **−84.0%** |

## 5. Gaps (in addition to 0370 §5 and 0370a §5)

| Family | Finding | Next step |
|---|---|---|
| shared rich-text payloads | refused; the host copies the payload and storage first (with the first text fragment only, a host defect for multi-fragment storages) | copy-on-write of payload and storage |
| Keynote slide-table cell lists (31 rows) | same host shape on slide tables | W1-C, over `storage_list` |
| no native rich-text cell in the corpus | every case promotes a plain cell or reads the result | a native fixture with rich-text cells (mac queue) |

## 6. Verification

On the merge of trunk `8022ffd507` (gate script
`.coord/iwa-w1a/gates3.sh`, sequential, `CARGO_BUILD_JOBS=6`), then the
placement follow-up:

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa -p litchi-pages -p litchi-keynote -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --features numbers --all-targets -- -D warnings` | 0 |
| `cargo test --all-features` for `litchi-numbers-wire`, `litchi-numbers`, `litchi-pages`, `litchi-keynote` | 0 each |
| `cargo test -p litchi-iwa --all-features` | 101 on the gate run (the two Pages cell parity cases above); after the follow-up the `pages_`, `keynote_slide_table`, `numbers_` parity filters pass (87 and 43 tests), and `litchi-numbers` 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers-wire -p litchi-numbers` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `cargo test -p litchi-ops --all-features` (lib and every integration test, in nine chunks) | 0 each |
