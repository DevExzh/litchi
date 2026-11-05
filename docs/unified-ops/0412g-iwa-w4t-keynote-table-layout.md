# 0412g — W4-T: Keynote slide-table resize, frame and reads

- Status: merge-ready on branch `uo2/iwa-w4t` after
  [0412f](0412f-iwa-w4t-conditional-highlight.md).
- Date: 2026-11-02.
- `performance_claim`: every route uses fewer instructions than its host
  route (§3).

Four of the five Keynote `table.lifecycle` rows are migrated:
`resize_slide_table`, `set_slide_table_geometry`, `slide_tables` and
`slide_table`. `add_slide_table` needs the from-scratch table graph. The
coordinator gave that graph, with Pages `body_table` and Keynote
`add_slide_table`, to W4-G, so the row stays host-only here.

## 1. Shared engine: `litchi_iwa_table::topology` layout

- `stage_resize(working, target, rows, columns)` is the Numbers
  trailing-edge resize (record 0370e: tiles, header buckets, UID map,
  stroke sidecar, the formula owner's extents, model counts), lifted from
  `litchi-numbers/src/package/table_resize.rs`. Numbers' `table::resize`
  family now calls it, and the Numbers resize tests and parity cases pass
  unchanged.
- `stage_uniform_dimensions(working, target, width, height)` gives every
  row the height `height / rows` and every column the width
  `width / columns`. It goes through the shared header-size codec
  (`table_dimension_codec::plan_header_storage_bucket_sizes`), the codec
  the host's `set_uniform_table_dimensions` uses, so the output is
  byte-identical.
- `frame` / `stage_frame` read and patch the table drawable's
  `TSD.GeometryArchive` field by field, as the host's
  `patch_drawable_geometry` does: present fields are patched in place and
  absent ones appended.
- New engine refusals: `InvalidSize`, `StoredCells` and `UnsupportedLayout`.

## 2. Facade

`litchi_keynote::slide::table::layout` adds
`Package::{edit_slide_table_layout, apply_slide_table_layout}` with
`Edit::{resize(rows, columns), set_frame(Geometry)}`. Each operation sizes
the rows and columns uniformly to fill the frame, as the host does. A frame
without a rotation keeps the table's rotation. The family has an
exact-source patch and inverse, and typed refusals.

The reads need no new API. `slide_tables` maps onto
`slide_drawable_list` (table frames), `slide_table_name` and
`slide_table_cells` (size). `slide_table` maps onto `slide_table_cells`
and `slide_table_merges`.

## 3. Parity and performance

| Test (`keynote_table_layout_cases.rs`) | Tier |
|---|---|
| `keynote_slide_table_resize_parity` (grow by 2 rows and 1 column) | A |
| `keynote_slide_table_frame_parity` (move and enlarge) | A |
| `keynote_slide_tables_read_parity` (names, sizes, frames on two native decks) | A |
| `keynote_slide_table_cells_and_merges_read_parity` | A |

Callgrind, release parity binary SHA-256
`93ca1b770e88969ec69c440a91474f11eda7ce074a28b5abe6d4ad4ed185daf1`, CPU 5,
ABBA:

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Keynote resize | 839,087,982 | 347,149,712 | 0.414 |
| Keynote frame | 739,972,237 | 345,316,835 | 0.467 |
| Numbers resize (`numbers_table_resize_cost_abba`, after the lift) | 157,596,452 | 62,407,588 | 0.396 |

Ledger: 4 Keynote rows migrate. Two native queue rows. The glossary
registers the family.

## 4. Verification

The gate script `(deleted-coord-artifact)` runs on the merge of trunk; its
result is reported in the hand-off.
