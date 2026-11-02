# 0415 — W4-P: image-fill replace/reset; scope status of the next package

- Status: in progress on branch `uo2/iwa-w4p` (trunk `d67a878026`
  merged). Item 2 landed. Items 1, 3 and 4 are not started; they are left
  for a continuation (§3).
- Date: 2026-11-02.
- Scope (coordinator): (1) `pages/creation.rs body_table` as a
  from-scratch table graph in `litchi-numbers-wire`; (2) replacing and
  resetting an existing image fill; (3) the Numbers drawable families
  `drawable.shape` (22), `drawable.image` (6), `drawable.movie` (6),
  `drawable.audio` (1); (4) `chart.lifecycle` on all three formats.
- `performance_claim`: none new (the image-fill route reuses the
  measured style publication; 0413b §4).

## 1. What landed (item 2)

| Commit | What |
|---|---|
| `d346840a0c` *(shared)* | `litchi_iwa_style::shape::property::ReleasingFills` (the fill property with a replaceable stored image fill) and `shape::stored_fill`, for a format family that releases an image fill's data reference and media itself |
| `c64fcd0c24` | Pages `drawable::style`: a fill set, a fill reset or a second image fill over an existing image fill releases the old style's data reference and its registered data-reference owner. A collapsed variation drops its owner. Data nothing else references loses its `DataInfo` and `Data/` member. These are the host's `remove_style_data_reference` and `remove_orphaned_image_asset`, on the same metadata transitions as the media family. Only two image fills staged in one edit stay `Refusal::ImageFill` |
| `231f208210` | parity `pages_body_shape_image_fill_release_parity` (solid fill, reset and new image over a host image fill; tier B; macOS-queued) |

No ledger row changes: the fill set and reset routes were already
migrated (record 0371); this closes their image-fill gap.

## 2. Gates for item 2

These passed (exit 0):

- strict clippy of litchi-iwa-style, litchi-pages and litchi-iwa (all
  targets, all features);
- the image-fill parity cases;
- the litchi-iwa-style and litchi-pages tests (811);
- rustdoc `-Dwarnings` for litchi-pages and litchi-iwa-style;
- the full glossary gate (73 profiles);
- the ledger check.

## 3. Not started (items 1, 3, 4)

- **body_table**: extend `litchi_numbers_wire::table_create` with a
  from-scratch table graph (drawable, model, tiles, header storages,
  UID maps, formula owner) matching the host Numbers builder's output.
  Use it from Pages `Builder::body_table` and the Numbers builder and
  table-add paths. Keep Numbers parity unchanged and macOS-queue the
  Pages route.
- **Numbers drawables**: port the Pages and Keynote shape lifecycle
  (`drawable::lifecycle` text-shape creation and duplication, presets and
  line segments on `litchi_iwa_style::shape_path`, image fills with the
  release transition above) and the shared caption seam to
  `litchi-numbers` `sheet::drawable`. Numbers already has drawable
  properties, media add, remove, duplicate and replace (0378, 0378a).
- **chart.lifecycle**: chart add, duplicate and remove for Numbers (7),
  Pages (8) and the shared row (1). Chart style stays with W4-C.
