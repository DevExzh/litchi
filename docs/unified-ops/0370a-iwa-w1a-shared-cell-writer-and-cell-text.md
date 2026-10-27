# 0370a — W1-A slice 2: shared cell writer parity, reopen adoption, cell text properties

- Status: landed on branch `uo2/iwa-w1a`, after slice 1
  ([0370](0370-iwa-w1a-numbers-tables.md), merged as `856fa248e2`); trunk
  merged after each coordinator merge.
- Date: 2026-10-26.
- `performance_claim`: the Numbers commit reopen is 21.4% cheaper on the
  topology route and 4.0% cheaper on the text-style route. The new
  cell-text route runs 48.4% below its host route (§4).

## 1. What landed

| Commit | Crates | What |
|---|---|---|
| `7b5e6ad781` | `litchi-numbers-wire`, `litchi-iwa` | shared cell writer: the last released key is handed back, and a new row's offset capacity matches; host defect fixed: a value clear dropped formatted cells (§2.1) |
| `d836e4d693` | `litchi-numbers` | every Numbers commit reopens through `Package::reopen_candidate` (`SourceCatalog::reopen`, W1-C) |
| `be95876ff9` | `litchi-numbers-wire`, `litchi-iwa` | shared cell writer: drops empty rows and bare zero-count headers; host defect fixed: zero-count row headers that held a height or style were dropped (§2.2) |
| `3f993f0164` | `litchi-numbers`, `litchi-iwa`, tools, docs | **new family `litchi_numbers::cell::text`** over W1-D's paragraph-style engine; 49 ledger rows migrated (§3) |
| this commit | docs | this record |

## 2. The shared table cell writer (`litchi_numbers_wire::table_write`)

Pages body tables (`edit_body_table_values`) and Keynote slide tables
(`edit_slide_drawables` cell writes) share this writer. W1-C and W1-B found
it diverging from the host after W1-B's host cell-writer fix (record 0371).

### 2.1 Clear and fill (`7b5e6ad781`)

- `list_release` now hands a released last key back to `next_list_id`, as
  both Numbers cell writers do.
- A new row's offset table takes the capacity of the widest existing row.
  Native rows keep spare slots, and the host sizes new rows the same way.
- **Host defect.** A value clear kept only commented cells. It dropped
  every other cell together with its style, text-style and format keys,
  so their list entries kept counting a cell that no longer existed. Now a
  cell whose metadata survives the clear is kept. Native Numbers keeps a
  cleared cell's formatting, and the Keynote and Pages leaf APIs document
  that rule ("the cell's style and format stay"). Axis removal and
  merge-anchor relocation still drop the cell, through the new
  `remove_attached_cell_in_package`. That keeps the host topology
  differential (`numbers_topology_formula_differential`) unchanged.

`keynote_slide_table_cell_write_corpus_parity` checks 46 writes, all tier B.
It replaces and clears every stored cell, and fills six empty cells with a
number and with text, in each slide table of `table-discovery.key`,
`slide-table-merges-native.key` and `slide-table-read-native.key`. One
refusal is listed: the Keynote leaf does not yet clear a formula cell
(W1-C).

### 2.2 Empty rows and headers (`be95876ff9`)

- `tile_set_cell` drops a row left without cells. It adjusts `numrows`
  under the tile's own convention (count or extent).
- `header_adjust` drops a header that counts no cell and carries no size,
  hidden state or style.
- **Host defect.** `update_row_header` dropped a zero-count row header even
  when it held the row's height, hidden state or styles. Now only a bare
  header goes.

`pages_body_table_cell_write_corpus_parity` checks 53 writes, all tier B.
It replaces and clears up to 24 stored cells and fills six empty cells per
body table, on five Pages fixtures plus W1-B's case: the host stores D2 of
`body-table-catalog-native.pages`, then both routes clear it. That case was
pinned in `uo2/iwa-w1b` `2f287d94e9`. The test lists the refusals for cells
of a table in a formula graph; those wait for W1-F's formula planner
(record 0376).

## 3. `litchi_numbers::cell::text`

The API is `Package::{edit_table_cell_text, apply_table_cell_text,
table_cell_text}`, with `Edit::set(At, Property)` and `Edit::reset(At,
PropertyKind)`. `Property` and `PropertyKind` are re-exported from
`litchi_iwa_text_wire::paragraph_style` (record 0374). Each staged change
runs the engine's `plan` over the cell's style. The style counts as
exclusive when its list entry has refcount 1, nothing but the list, the
stylesheet and the style references it, and no style names it as parent.
The family then publishes the plan:

- **Rewrite**: the variation's payload is replaced in place, with its
  managed field infos synchronised.
- **Derive**: a child variation is stored beside its stylesheet and
  registered there. A list entry names it, the cell points at the entry,
  and an empty cell takes a style-only cell.
- **Collapse**: the cell points back at the parent, with no key when the
  parent is the section's base style. The entry is released, and the
  variation is removed with its stylesheet registration and UUID. A new
  `Transition::release_suffix` lowers the identifier watermark when the
  variation was its suffix, as the host does.

**Host defect.** A reset back to the section base style left the style
list's external reference to the parent, although no entry named it any
more.

Parity runs on B2 of `basic.numbers` for the 17 host properties: text
style, font, decorations, colour, capitalization, script, baseline shift,
character spacing, ligatures, outline, shadow, background, alignment, line
spacing, paragraph spacing, indents and tab stops. Each property's sample
(W1-D's `text_style_cases::sample`) is set, then reset on the host's
output, then the next property is set on that output so the variation is
rewritten in place. All 51 routes pass tier B
(`numbers_cell_text_property_parity`). The 17 reads are equal on 595 corpus
cells (`numbers_cell_text_property_read_parity`). The ledger migrates 49
rows: 16 setters, 17 resets and 16 reads. `set_table_cell_text_style` and
`table_cell_text_style` stay on `edit_table_cell_text_style` (slice 1). The
family is registered in the `numbers-minimal` and `numbers-full` glossary
profiles (33 families) and in `docs/API_GLOSSARY.md`.

## 4. Performance

Callgrind on the release parity binary, CPU 3 pinned. Each probe runs the
host route, the leaf route twice, then the host route again
(`cost_probes.rs`).

| Route | Before | After | Δ |
|---|---:|---:|---:|
| leaf topology insertion (reopen adoption) | 528,110,642 | 415,246,013 | −21.4% |
| leaf whole-cell text style (reopen adoption) | 336,641,153 | 323,200,560 | −4.0% |

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| cell text: alignment set, then reset, `basic.numbers` B2 | 626,931,746 | 323,399,324 | **−48.4%** |

## 5. Gaps (in addition to 0370 §5)

| Family | Finding | Next step |
|---|---|---|
| cell paragraph lists (30 rows) | the host promotes the cell to rich text (a `TSWP.StorageArchive` in a rich-text payload, `cell_paragraph_list.rs`, 2.6k lines), then runs the text list editor on that storage | port rich-text promotion into the leaf, then publish W1-D's `list` engine on the cell's storage |
| generic `DataFormat` and numeral system (8 rows) | the host delegates the eligible formats to the leaf's dedicated families; numeral system and cross-family changes go through its compatibility writer, which has no leaf counterpart | one format family that dispatches to the dedicated writers, plus a numeral-system writer |
| cell comments, conditional highlights, lifecycle, removal, duplicate, Builder | unchanged from 0370 §5 | — |

## 6. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa -p litchi-pages -p litchi-keynote --all-targets --all-features -- -D warnings` | 0 |
| `cargo test --all-features --no-fail-fast` for `-p litchi-numbers-wire`, `litchi-numbers`, `litchi-iwa` (every target, including `parity` and the topology formula differential), `litchi-pages` and `litchi-keynote` | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal` / `numbers-full` | 0 / 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
