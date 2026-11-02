# 0413b — W4-P continued: shared preset paths, Pages shape lifecycle, presets and image fills

- Status: the continuation of [0413](0413-iwa-w4p-pages-text-boxes-shapes-comments.md)
  on branch `uo2/iwa-w4p` (trunk merged before each hand-off). The preset
  and lifecycle slices are merged. The image-fill slice is ready. One row
  stays host-only (§5).
- Date: 2026-11-02.
- Scope: the coordinator's four items after 0413:
  1. the shared preset path builder;
  2. the adds, duplicates and `body_shapes`;
  3. the 0371 image-fill gap;
  4. the `pages/creation.rs body_table` row.

  A fifth item came later: removing a commented drawable must clear its
  comment first instead of refusing.
- `performance_claim`: every new route is 68–81% cheaper than its host
  route under callgrind ABBA (§4).

## 1. What landed

| Commit | What |
|---|---|
| `819712809b` *(shared)* | `litchi_iwa_style::shape_path::{path_source, line_path_source, preset, with_preset}`. These are the host's `shape_path_source`, `shape_preset` and `set_shape_preset` over the common `Preset`. The Numbers sheet-drawable author and the Keynote slide-drawable author and preset reader now use it, and their identical copies are deleted |
| `8d1c8ccec1` | Pages `drawable::properties::{Snapshot::preset, Edit::set_preset, Preset}` for shapes and text boxes |
| `19dc47cd1c` | `drawable::lifecycle::Edit::{add_text_box, add_shape, add_line}` build the host's text-shape graph under fresh identifiers in the host order (drawable, caption, title, storage, attachment): the shape, two stand-ins, a storage inheriting the body's stylesheet, first paragraph and list styles and language, and the body attachment. They also add the U+FFFC and attachment entry, the z-order entry (creating the z-order when absent), the UUIDs, the style's external reference and the watermark. New refusals: `InvalidGeometry`, `MissingThemeStyle` |
| (with `9724c9a34f`) | `Edit::remove` clears a commented drawable's thread through `drawable::comment` before removing it, as the host does. The patch spans both steps |
| `428c32a13c` *(host fix)* | `Draft::duplicate_body_shape` compared a copied line's segment with the source's unmoved segment, so it refused every line duplication. It now expects the 12 pt offset |
| `a01e46f6c8` | `Edit::duplicate` covers shapes, lines and text boxes: the host graph orders, shape-info references remapped, the nested frame moved 12 pt. `Edit::duplicate_text_box` sets the copy's text through `drawable::text` |
| `3292d9fca1` | `drawable::style::Edit::set_image_fill(bytes, file_name, ImageTechnique, tint)`. The bytes become a new `Data/` member and `DataInfo` (next data identifier, `stem-identifier.extension`). The fill goes through the shared plan layer (`shape::{target, plan_set, apply}`). The shape's style variation takes the data reference and becomes the registered data-reference owner. This is the media family's data-owner transition and Keynote's image-fill path |

## 2. Ledger

24 Pages rows migrate in this continuation:

| Rows | Host routes | Destination | Tier |
|---:|---|---|---|
| 2 | `{,set_}body_shape_preset` | `drawable::properties` | A (byte-identical) |
| 7 | `add_text_box`, `add_body_{rectangle,shape,shape_with_fill,line,line_with_endpoints,line_with_style}` | `drawable::lifecycle::Edit::{add_text_box, add_shape, add_line}`; the combined routes are a creation followed by `drawable::style` | B |
| 2 | `duplicate_text_box`, `duplicate_body_shape` | `Edit::{duplicate_text_box, duplicate}` | B |
| 1 | `body_shapes` | `Package::body_drawable_properties` with `body_drawable_text` and `body_drawable_style` | A |
| 1 | `set_body_shape_image_fill` | `drawable::style::Edit::set_image_fill` | B |

Every route has a W4-P row in `0363-mac-queue.md`, as do the commented
removal and the host fix. With 0413, W4-P has migrated 203 rows and
retired one. `python3 tools/iwa_exit_ledger.py --check` passes.

## 3. Parity findings

- **Identifier order.** Creations and copies follow the host's graph
  order, so new identifiers equal the host's. Only the documented UUID and
  save-token policies make them Tier B.
- **Combined host routes.** The host's with-fill, with-endpoints and
  with-style adds call its own style setters after creating. The leaf
  composes the creation and `drawable::style`, and the bytes agree under
  Tier B.
- **Line duplication.** This was a host defect (§1) and is fixed and
  queued.
- **Image fills.** An existing image fill cannot be replaced or reset
  here (`Refusal::ImageFill`). Releasing the old asset is the reset path
  of the host's `reset_shape_fill`, which no migrated row needs.

## 4. Performance

Callgrind on a release build with debug info, CPU 3, ABBA, full routes
(open, edit, commit, serialize). Probes are in
`pages_text_box_lifecycle_cases.rs`.

| Route | Host Ir | Leaf Ir | Δ |
|---|---:|---:|---:|
| text-box add | 350,908,189 | 111,594,954 | −68.2% |
| shape duplicate | 53,332,338 | 10,438,028 | −80.4% |
| preset set | 19,468,099 | 3,682,653 | −81.1% |
| shape image fill | 35,008,402 | 11,229,792 | −67.9% |

## 5. `pages/creation.rs body_table` (left host-only)

W1-B's `table::creation` (record 0371) clones an existing body table as its
template, so it covers tables added to documents that already hold one. On
a table-less document it is the typed refusal `Refusal::NoTemplate`. The
host's `Builder::body_table` instead synthesizes the first table graph with
its Numbers builder and transplants the objects. A leaf crate cannot do
that: `litchi-pages → litchi-numbers` is a leaf-to-leaf edge outside the
policy (D4).

Retiring the row would lose a capability, so it stays host-only. The
leaf route needs one of:

- a small bundled template with a 1×1 table (a `blank-table.pages` asset
  built once by the host builder, with provenance recorded) that
  `Builder::body_table` starts from, resizing and renaming its table
  through `table::size` and `table::name` and adding further tables
  through `table::creation`;
- a from-scratch table graph builder in `litchi-numbers-wire` (the shared
  table engine), which both builders would use.

The first is the smaller change. Its output differs structurally from the
host's synthesized graph, so it would be macOS-queued, not Tier A.

## 6. Verification

Full gate on `a456664b12` (trunk `e1b061796d` merged), every row exit 0:
fmt; strict clippy of litchi-pages, -numbers, -keynote, -iwa-style,
-iwa-text-wire, -iwa-table, -iwa, -ops, litchi and litchi-py (all features)
and of litchi-ops with no features, `pages`, `keynote` and `numbers`;
tests of litchi-pages (768), -numbers, -keynote, -iwa-style, -iwa-text-wire,
-iwa-table, litchi-iwa (2,744, parity included), litchi and litchi-ops (lib,
doc, nine test groups); rustdoc `-Dwarnings`; the glossary gate (73
profiles); crate boundaries; `non_iwork_gate`; the tool unit tests; the
ledger, capability-delta, corpus, bench-matrix (regenerated for this
record) and feature-matrix checks; pyo3; perf-baseline; every fuzz lock
offline and locked. Both pre-commit size checks printed
nothing for every commit.
