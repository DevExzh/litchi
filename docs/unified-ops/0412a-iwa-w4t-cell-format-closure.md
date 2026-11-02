# 0412a — W4-T: closing the Pages and Keynote cell-format rows

- Status: merge-ready on branch `uo2/iwa-w4t` after record
  [0412](0412-iwa-w4t-cell-rich-text.md) (slice 4 merged as `aef9c70ab4`).
- Date: 2026-11-02.
- `performance_claim`: every new route uses fewer instructions than its
  host route except Keynote Automatic to Date and Time (+2.15%, within the
  5% bar; §4).

Every `table.cell_format` row of Pages and Keynote is now migrated (29 more:
14 Pages, 15 Keynote). The coordinator ruled that none of the 0412 §5
reasons is terminal; each is closed below.

## 1. Pop-Up Menu: the shared publication

- `litchi_iwa_table::cell_format::pop_up_publish::publish` is the Numbers
  Pop-Up Menu transaction's publication, lifted: native graph rewrite,
  metadata registry transition, preview policy, exact reassembly, leaf
  reopen (a closure), package locality (`verify_locality`) and readback.
  The ZIP-record locality checks move with it
  (`cell_format::locality`); Numbers' table-header helpers re-export them.
- Numbers' `table_cell_pop_up_menu` now calls it (previews deleted, keys
  reused, model UUID registered: its behaviour is unchanged; every Numbers
  test passes) and its `verify_package_locality_for_members` delegates to
  `verify_locality`.
- Pages and Keynote call it with the host policy their parity pins:
  previews kept, list keys appended, no UUID or external edge registered for
  a menu model (`register_model_uuid = false`). The registry transition then
  only advances the watermark and save tokens; an unrecorded model edge is
  skipped on removal (`RegistryFacts::recorded_external_removals`).
- Native Pages and Keynote tables broke three assumptions of the Numbers
  engine, each fixed in the engine:
  - the BNC census counted only control cells and refused format-only cells;
    it now counts every format reference (primary and secondary) and only the
    selected table's tiles (a component holds several tables' tiles);
  - a cell naming an implicit Number entry (decimal kind, generic identifier,
    no explicit flag, no control) gives that entry up when it becomes a menu,
    in the native list rewrite and in `litchi_numbers_wire::popup_menu`'s
    BNC transition;
  - a cleared menu's last CellSpec key goes back to `next_list_id` under the
    append policy.
- `Refusal::UnsupportedFormat` of the Pages and Keynote format families is
  deleted: nothing refuses with it any more.

## 2. Custom formats

The custom-format registry is the document root's `TSA.DocumentArchive`
`custom_format_list` (field 12) in all three formats; the engine knew only
the Numbers document. `display::document_layout` now resolves the Numbers
(`TN.DocumentArchive`, super 8, direct field 9), Keynote
(`KN.DocumentArchive`, super 3, direct field 4) and Pages
(`TP.DocumentArchive`, message type 10000, super 15, direct field 41) roots.
The Pop-Up Menu control reader decoded every format payload with the control
codec, which refuses custom (270–272) payloads; for a cell without a control
spec it now reads only the format type it checks. No fixture had to be
authored: the native fixtures already carry an empty registry, which the
host and leaf writers fill.

The host mints random registry UUIDs, so the write parity is a read
differential: both outputs reopen, host and leaf reads of both agree, the
leaf inverse restores the source, and both routes' resets on both outputs
read Automatic.

## 3. Date and Time, Duration, Currency and Text on native cells

- **Reads (a leaf defect).** The control reader refused every cell of kind
  3 (Date and Time) or 4 (Duration) as an invalid source before the display
  readers saw it. A cell of those kinds without a control spec now holds no
  control.
- **Writes.** On a native cell naming an implicit Number entry, the host's
  compatibility writer converts: Currency (type 10, explicit tuple), Date
  and Time (type 9), Duration (type 7, the value in seconds) and, on a text
  cell, Text (kind 5, explicit text flag), each with a freshly interned
  entry and the implicit entry released. The engine's `rewrite_display_format`
  admits that conversion (`converting`) and builds the cell with
  `BncCell::set_data_format_identifier`. A cleared Date and Time cell holding
  a plain number returns to type 2 (`clear_explicit_format`), as the host's
  reset does. The macOS queue carries the native check of this host shape.
- **Keynote Text fixture.** Both cell writers refuse
  `slide-table-read-native.key`'s table: its formula owner caches a
  non-scalar value the host cannot refresh, and the leaf slide-drawable
  writer refuses formula tables (`TableFormulas`). That is a formula refresh
  limitation, not an admission defect of these formats; the Text cases use
  `table-discovery.key` B2, written by the host's cell writer.

## 4. Parity and performance

| Test (`pages_body_table_` / `keynote_slide_table_` prefix) | Tier |
|---|---|
| `cell_date_duration_format_set_parity`, `_reset_parity` | B |
| `cell_date_duration_format_read_parity` | A |
| `cell_data_format_set_parity` (now Number, Currency, Percentage, Scientific, Fraction, Date and Time, Duration) | B |
| `cell_custom_format_parity` | B (read differential) |
| `cell_custom_format_read_parity` | A |
| `cell_pop_up_menu_format_parity` (set, replace, reset) | B |
| `cell_pop_up_menu_format_read_parity` | A |
| `cell_text_format_parity` (set, reset) | B |
| `cell_text_format_read_parity` | A |

Callgrind, release parity binary SHA-256
`b1f47aa4b3635111851fc7dbf0521d05f5631f6584fe9595ff9f88ee51d7f6eb`, CPU 5,
`pages_cost_probes::{probe_host,probe_leaf}`, ABBA:

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Pages Pop-Up Menu set | 3,417,371,211 | 332,631,651 | 0.097 |
| Pages Automatic to Date and Time | 3,356,004,724 | 398,452,259 | 0.119 |
| Keynote Pop-Up Menu set | 852,272,037 | 550,412,233 | 0.646 |
| Keynote Automatic to Date and Time | 785,243,616 | 802,114,728 | 1.021 |
| Pages Checkbox (re-run) | 3,364,061,363 | 284,340,524 | 0.085 |
| Keynote Checkbox (re-run) | 794,068,759 | 642,726,461 | 0.809 |

Ledger: 14 Pages and 15 Keynote rows migrate; ten native queue rows.

## 5. Verification

See the hand-off message of this slice for the gate results on the merged
tip.
