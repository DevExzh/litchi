# litchi-iwa

Legacy migration host for Apple iWork archive work on `.pages`, `.numbers`,
and `.key` files.

## Overview

`litchi-iwa` reads Apple iWork bundles using their IWA (iWork Archive) layout:
a ZIP container holding Snappy-compressed, protobuf-encoded object streams
along with media assets and metadata. It is the legacy migration host, not the
supported format facade. Its remaining public surface is for compatibility
adapters and editor capabilities that have not yet moved to a concrete format
crate.
This package is intentionally unpublished (`publish = false`); use it only as
a workspace or source-checkout dependency while migrating existing callers.

New semantic code belongs in `litchi-pages`, `litchi-numbers`, or
`litchi-keynote`; use `litchi::iwork` for a supported cross-format snapshot.
In particular, ordinary Keynote slide title, body, and speaker-notes reads or
edits must use `litchi-keynote::Package` and `SlideSelector`. Presentation-wide
dimensions and show playback flags use `litchi_keynote::show`; playback mode
and volume for an existing soundtrack use `litchi_keynote::soundtrack`, not
`litchi_iwa::Document` or `KeynoteEditor`. The concrete package keeps native
identifiers and raw records private, validates semantic ownership, and creates
exact-source checked commits.

```rust,no_run
use litchi_keynote::{Package, SlideSelector};

let package = Package::open("input.key")?;
let mut edit = package.edit_slide_body(SlideSelector::index(0))?;
edit.set("Updated body")?;
let commit = edit.commit()?;
let mut output = Vec::new();
commit.package().write_to(&mut output)?;
assert!(!output.is_empty());
# Ok::<(), Box<dyn std::error::Error>>(())
```

To publish a changed package, use the focused
`litchi-keynote/examples/edit_slide_text.rs` or
`litchi-keynote/examples/edit_show_settings.rs` workflow. Each requires a
distinct, new output path, writes with `Package::write_to` through a sibling
temporary file, synchronizes it, and uses no-clobber publication; do not write
directly to an existing path.

Presentation settings use the same immutable package chain and no native
identifiers:

```rust,no_run
use litchi_keynote::{
    Package,
    show::{Mode, Size},
};

let package = Package::open("input.key")?;
let before = package.show_settings()?;
let mut settings = before;
settings.set_size(Size::new(1920.0, 1080.0)?);
settings.set_mode(Some(Mode::SelfPlaying))?;
settings.set_loop_presentation(Some(true));

let commit = package.edit_show_settings()?.set(settings).commit()?;
assert_eq!(
    commit.package().show_settings()?.mode(),
    Some(Mode::SelfPlaying),
);
let restored = commit
    .package()
    .apply_show_settings(&commit.patch().inverse())?;
assert_eq!(restored.package().show_settings()?, before);
# Ok::<(), Box<dyn std::error::Error>>(())
```

## Usage

`litchi-iwa` is intentionally unpublished and is not a crates.io dependency.
For a local source checkout, depend on it by path while migrating existing
callers:

```toml
[dependencies]
litchi-iwa = { path = "../litchi-iwa" }
```

The unified `Document` API below is retained for legacy compatibility
inspection. It is not the entry point for new format-specific semantic code.

```rust
use litchi_iwa::Document;

let doc = Document::open("document.pages")?;
let text = doc.text()?;
let stats = doc.stats()?;
println!("objects: {}", stats.total_objects);

# Ok::<(), litchi_iwa::Error>(())
```

## Features

- Legacy-compatible inspection of Pages, Numbers, and Keynote bundles from a
  path or in-memory bytes
- Build independent Pages, Numbers, and Keynote packages from typed IWA objects
  with no bundled template
- Snappy decompression and protobuf decoding of `.iwa` streams
- Text extraction across all iWork applications
- Metadata-backed media discovery, extraction, replacement, and guarded cleanup
- Typed cross-suite image, movie, and audio-property read/write for hyperlinks,
  locking, aspect-ratio locking, and accessibility descriptions with lossless
  unknown-field preservation
- Metadata-preserving IWA object/message create, read, update, and delete operations
- Snappy IWA serialization and deterministic package rewriting
- Transactional package-entry and IWA-component updates with atomic saves
- Legacy nested `Index.zip` bundle import with byte-preserved assets and
  explicit password-protected document rejection
- Legacy host editors for Numbers sheets/tables/cells/formulas, Pages
  body/text-box text, and unmigrated Keynote slide graphs and arbitrary text
  boxes. Selector-first Pages header/footer text now belongs to
  `litchi-pages::Package`; selector-first Keynote title, body, and
  speaker-notes text and presentation-settings transactions are owned by
  `litchi-keynote::Package`.
- `litchi-pages::Package` selector-first section-text reads and, for rooted
  exact sources with one unambiguous native body storage,
  set/clear/UTF-16-span transactions with reversible patches. Checked
  `TextPosition` and insertion-capable `TextSpan` values keep byte offsets,
  native object identifiers, and protobuf records out of the public API.
- `litchi-pages::Package` selector-first aggregate section-settings reads and
  exact-source transactions. The semantic value preserves optional name,
  Boolean, and pagination presence while `section::settings` owns the direct
  edit, patch, diagnostics, error, limit, dependency, and path types.
- Native-style ordinary shape duplication across Pages, Numbers, and Keynote
  with independent rich-text storage, fresh UUID mappings, preserved opaque
  fields, and app-specific selection offsets
- Native Pages, Numbers, and Keynote chart CRUD, including source-built inline
  data charts, typed native caption CRUD, and duplicate operations with fresh
  private graphs, preserved editable data, theme-preset registration, UUIDs,
  and native placement offsets
- Selector-first Keynote chart-legend visibility through
  `litchi-keynote::Package`, with semantic slide/chart selectors, exact-source
  no-op and inverse patches, and private native graph identifiers
- Typed copy-on-write text-box paragraph alignment, native line-spacing modes,
  atomic before/after spacing, first-line/left/right indentation, and ordered
  left/center/right/decimal tab stops with leaders across Pages, Numbers, and Keynote
- Typed BCP 47 text-language run CRUD at validated UTF-16 scalar boundaries,
  including automatic-language sentinels and lossless boundary deletion
- Native cross-suite text-hyperlink CRUD with typed nonempty UTF-16 ranges,
  lossless web, mail, and Keynote navigation targets, and owned-object cleanup
- Native cross-suite Date & Time smart-field CRUD with typed native date/time
  pattern strings, locale identifiers, formatter styles, refresh plans, and
  Apple-reference instants; pattern grammar is not validated
- Native Pages page-number/page-count attachment CRUD with typed kinds, exact
  U+FFFC placement, lossless number metadata, and header/footer support
- Native Pages body-bookmark CRUD with typed names, lossless visibility values,
  strict UTF-16 ranges, and owned bookmark-field cleanup
- Native cross-suite plain-text highlight CRUD with typed nonempty UTF-16 ranges,
  lossless table mutation and owned annotation cleanup
- Native cross-suite ranged text-comment and ordered direct-reply CRUD with
  nonempty typed bodies, stable IDs and metadata, and scratch-package creation
- `litchi-pages::Package` owns lossless Pages page-layout transactions for
  dimensions, margins, scale, orientation, and vertical body layout
- Typed Numbers table header/footer counts, freeze state, and repeating-header
  settings with lossless optional-field presence
- Typed Numbers full-table sort-rule configuration CRUD with lossless native
  rule and reference-tracker preservation
- Lossless Numbers table-title visibility and title-outline settings
- Typed Keynote theme-layout discovery and fresh empty-slide creation with native
  component registration, speaker notes, storage-less slide-number placeholders,
  and transactional insertion
- Source-free and wire-preserving Keynote per-slide number visibility with native
  placeholder ownership and z-order invariants
- Native Keynote build-in/build-out effect CRUD with typed On Click / After Transition /
  With Previous / After Previous timing, typed Rotate / Scale / Opacity / Move actions,
  editable Bézier motion paths and custom timing curves, typed Blink / Bounce / Flip / Jiggle / Pop / Pulse
  emphasis actions, typed Keyboard / Shimmer / Skid / Swoosh / Trace build-in/build-out
  effects with native direction models, validated raw CRUD for unmapped native build parameters,
  component UUIDs, and slide-node cache maintenance. Focused `litchi-keynote` owns the
  bounded existing-build playback-order transaction; this host retains effect, timing,
  and add/update/remove compatibility only.
- `litchi-keynote::transition` selector-first modern slide-transition reads
  and exact-source set/clear transactions with reversible patches. A private
  Buffa lazy view projects known fields, while validated raw records remain the
  preservation authority.
- Typed direct-drawable comment CRUD with Pages document-reachability,
  Numbers sheet-ownership, and Keynote slide-ownership guards
- Native Numbers cell-comment and direct-reply CRUD with table-list refcounts,
  copy-on-write threads, annotation authors, dates, and UUIDs

## Legacy editing

The examples in this section document remaining migration-host APIs. They are
appropriate when a workflow explicitly needs an unmigrated editor capability
or native/archive compatibility behavior. Do not use them as a substitute for
the concrete Pages, Numbers, or Keynote package APIs.

Native charts can be created directly from typed `ChartData` with
`add_body_chart`, `add_sheet_chart`, or `add_slide_chart`; no source table or
template package is required. Their corresponding `duplicate_body_chart`,
`duplicate_sheet_chart`, and `duplicate_slide_chart` methods retain the source
data and opaque chart fields while giving the new chart independent inline data,
private styles, preset registration, UUIDs, ownership, and native placement.
See `create_*_chart` and `duplicate_*_chart` in `examples/` for runnable
file-to-file workflows.

Charts expose their native title through `body_chart_title`,
`sheet_chart_title`, and `slide_chart_title`. `set_*_chart_title` updates the
same `Chart Options > Title` state and text that Pages, Numbers, and Keynote
save; `remove_*_chart_title` returns whether a title was visible. Chart titles
remain independent through duplicate, delete, and package round-trip
operations.

The native Axis formatter is available through the typed
`Axis::{Category, Value}` selector. Use `body_chart_axis_title`,
`sheet_chart_axis_title`, or `slide_chart_axis_title` to read an axis name,
then `set_*_chart_axis_title` or `remove_*_chart_axis_title` to update the
same `Axis > Category (X) / Value (Y) > Axis Name` controls that Pages,
Numbers, and Keynote save. `Value` selects the primary value-axis object;
titles remain independent through duplicate, delete, and package round-trip
operations.

The value-axis `Min` and `Max` fields are represented without sentinel values
by `Bounds` and `Bound`. Pages and Numbers still expose their corresponding
legacy migration-host bounds, steps, and scale methods while those format
owners are being migrated. Keynote value-axis settings are now owned by the
selector-first `litchi_keynote::Package` aggregate API:
`slide_chart_value_axis_settings`,
`edit_slide_chart_value_axis_settings`, and
`apply_slide_chart_value_axis_settings`. Select a slide with
`SlideSelector` and a chart with `ChartSelector`, then update one
`ValueAxisSettings` value atomically through the edit's `set`, `set_bounds`,
`set_steps`, or `set_scale` methods. The focused owner validates graph
ownership, preserves unknown native fields, and keeps exact-source inverse
patches; it does not expose Keynote drawable IDs. Each bound remains
independently optional (`None` means the app's `Auto` value), and invalid
non-finite or inverted ranges are rejected before package mutation.

Axis-line visibility is likewise typed through
`body_chart_axis_line_visible`, `sheet_chart_axis_line_visible`, and
`slide_chart_axis_line_visible`; use `set_*_chart_axis_line_visible` to
update the same `Axis > Category (X) / Value (Y) > Axis Line` switch that all
three apps save. Source-built charts default both primary axis lines to
visible, and each axis remains independently configurable through duplicate
and package round-trip operations.

Major-gridline visibility uses the parallel
`body_chart_axis_major_gridlines_visible`,
`sheet_chart_axis_major_gridlines_visible`, and
`slide_chart_axis_major_gridlines_visible` APIs, with
`set_*_chart_axis_major_gridlines_visible` updating the native
`Axis > Category (X) / Value (Y) > Gridlines / Major Gridlines` state. New
column charts match iWork's native defaults: category-axis major gridlines are
hidden and value-axis major gridlines are visible. Each axis remains independent
through duplicate and package round-trip operations.

Minor-gridline visibility is independently typed through
`body_chart_axis_minor_gridlines_visible`,
`sheet_chart_axis_minor_gridlines_visible`, and
`slide_chart_axis_minor_gridlines_visible`; use
`set_*_chart_axis_minor_gridlines_visible` for the native
`Axis > Category (X) / Value (Y) > Minor Gridlines` control. New column charts
start with minor gridlines hidden on both primary axes, and duplicate and
package round-trip operations keep each axis independent.

Chart legend visibility remains a migration-host operation for Pages and
Numbers through `body_chart_legend_visible` / `set_body_chart_legend_visible`
and `sheet_chart_legend_visible` / `set_sheet_chart_legend_visible`. For
Keynote, use the focused selector-first package transaction instead:

```rust,no_run
use litchi_keynote::Package;

let package = Package::open("input.key")?;
let before = package.slide_chart_legend_visible("Charts", "Revenue")?;
let commit = package
    .edit_slide_chart_legend("Charts", "Revenue")?
    .set(!before)
    .commit()?;
assert_eq!(
    commit.package().slide_chart_legend_visible("Charts", "Revenue")?,
    !before,
);
let restored = commit
    .package()
    .apply_slide_chart_legend(&commit.patch().inverse())?;
assert_eq!(
    restored.package().slide_chart_legend_visible("Charts", "Revenue")?,
    before,
);
# Ok::<(), Box<dyn std::error::Error>>(())
```

The focused Keynote owner keeps native chart graph identifiers and raw records
private, preserves unknown fields, and validates exact-source no-op/change,
candidate readback, and inverse behavior.

Charts also expose their native generic Object Caption control through
`body_chart_caption`, `sheet_chart_caption`, and `slide_chart_caption`.
`set_*_chart_caption` creates or updates the real caption text graph, while
`remove_*_chart_caption` returns whether a caption was present. Captions remain
independent through duplicate, delete, and package round-trip operations.

Source-built and existing file-backed images also expose typed shared drawable
properties through `body_image_properties` / `set_body_image_properties`,
`sheet_image_properties` / `set_sheet_image_properties`, and
`slide_image_properties` / `set_slide_image_properties`. This makes image alt
text, hyperlinks, and lock state editable without raw protobuf mutation; see
the `create_*_image` examples.

The same image APIs expose the native basic Image-inspector controls through
typed `ImageAdjustments`: exposure and saturation use checked normalized values
from `-1.0` to `1.0` (`0.25` is `25%`), while `ImageEnhancement` models the
automatic Enhance switch. The setters preserve all unmapped advanced native
adjustments and opaque wire fields.

Source-built and existing direct drawables in the compatibility host expose
native Arrange stacking through `sheet_drawable_order` and
`slide_drawable_order`. Each list runs back-to-front; its setter requires an
exact permutation, while `move_*_drawable` accepts typed
`DrawableLayerMove::{ToBack, Backward, Forward, ToFront}` commands. See the
`create_numbers_stacked_shapes` and `create_keynote_stacked_shapes` examples
for complete scratch-file workflows. Pages body drawable stacking is owned by
the focused `litchi-pages` package and uses source-bound handles rather than
the legacy raw-ID editor API.

Ordinary source-built and existing shapes expose the native Flip buttons via
`DrawableFlipAxis::{Horizontal, Vertical}` and `flip_body_shape`,
`flip_sheet_shape`, or `flip_slide_shape`. Each operation updates native
geometry while retaining unrelated fields; see `create_*_flipped_shape` for
scratch-file examples.

File-backed images use the same typed command through `flip_body_image`,
`flip_sheet_image`, or `flip_slide_image`, preserving their embedded asset,
adjustments, and metadata; see `create_*_flipped_image` for scratch-file
examples.

Images with native original-size metadata can also restore just their displayed
dimensions through `restore_body_image_original_size`,
`restore_sheet_image_original_size`, or `restore_slide_image_original_size`.
Those operations retain the current position and transform while returning an
error for media that has no original dimensions; see
`create_*_original_size_image` for scratch-file examples.

Pages, Numbers, and Keynote image APIs also expose their native title/caption
controls. `*_image_title_caption` returns a shared `DrawableTitleCaption`;
`set_*_image_title` and `set_*_image_caption` create or update the typed native
text graphs, while the corresponding `remove_*` calls return whether a value
was present. See `create_*_image_caption` for complete source-free examples.

Ordinary body, sheet, and slide shapes expose the same native controls through
`*_shape_title_caption`, `set_*_shape_title`, `set_*_shape_caption`, and their
matching `remove_*` methods. Shape labels remain independent through duplicate,
delete, and package round-trip operations; the `create_*_shape` examples build
them from scratch.

File-backed movies in this migration host retain media, geometry, playback,
and other format-specific compatibility controls. Body and sheet movie labels
continue to use the corresponding `*_movie_title_caption`,
`set_*_movie_title`, `set_*_movie_caption`, and `remove_*` methods. Keynote
slide movie title/caption CRUD is owned by the focused
`litchi_keynote::Package` methods
`{slide_movie_title,edit_slide_movie_title,apply_slide_movie_title,slide_movie_caption,edit_slide_movie_caption,apply_slide_movie_caption}`
with `SlideSelector` and `MovieSelector`; the migration host no longer exposes
wrappers for those operations. See `create_keynote_movie.rs` for the focused
package workflow. Movie labels remain independent through duplicate, delete,
and package round-trip operations.

Body and sheet movies likewise expose `flip_body_movie` and `flip_sheet_movie`,
preserving their video and poster assets, playback settings, and metadata; see
`create_*_flipped_movie` for scratch-file examples. Keynote slide-movie
reflection is owned by the focused `litchi-keynote` geometry transaction; the
host's typed selector bridge delegates to that owner and retains no independent
flip implementation.

Movies with native original-size metadata can also restore just their displayed
dimensions through `restore_body_movie_original_size`,
or `restore_sheet_movie_original_size`. Keynote slide movie position,
displayed-size, and supported transform edits, including original-size
restoration, use the selector-first
`litchi_keynote::Package::{slide_movie_geometry,edit_slide_movie_geometry,
apply_slide_movie_geometry}` API; native angle and reflection flags are part of
that focused transaction. The compatibility host retains only its delegating
selector bridge plus unrelated movie creation/removal, media, offset, property,
and graph operations. Those
operations retain the current position and transform while returning an error
for media that has no original dimensions; see
`create_*_original_size_movie` for scratch-file examples.

### Create Pages documents from scratch

```rust
use litchi_iwa::pages::{PagesEditor, PagesImageOptions};

let mut pages = PagesEditor::builder()
    .body_text("Created entirely by litchi-iwa")
    .language("en")
    .locale("en_US")
    .build()?;
pages.set_body_text("Created and then updated through the same typed API")?;
let bookmark = pages.add_body_bookmark(
    litchi_iwa::text::TextRange::from_utf16_indexes(0, 7)?,
    litchi_iwa::text::TextBookmarkSettings::new().with_name(
        litchi_iwa::text::TextBookmarkName::new("Created")?,
    ),
)?;
pages.update_body_bookmark(
    bookmark.id,
    litchi_iwa::text::TextRange::from_utf16_indexes(8, 11)?,
    litchi_iwa::text::TextBookmarkSettings::new(),
)?;
pages.remove_body_bookmark(bookmark.id)?;
pages.save("created.pages")?;
# Ok::<(), litchi_iwa::Error>(())
```

Scratch documents can include an independently writable native table from the
first build. Its name, dimensions, and cells remain editable after reopening;
shrinking rejects any operation that would discard stored cells:

```rust
use litchi_iwa::numbers::CellValue;
use litchi_iwa::pages::PagesEditor;
use litchi_pages::{BodyTableSelector, Package as PagesPackage};

let mut pages = PagesEditor::create_with_text("Quarterly revenue\n")?;
let first_anchor = pages.body_text()?.encode_utf16().count();
let table = pages.add_table(first_anchor, "Revenue", 4, 3)?;
pages.set_table_cell(
    table.model_object_id,
    0,
    0,
    CellValue::Text("Quarter".to_owned()),
)?;
let selector = BodyTableSelector::index(0);
let focused = PagesPackage::from_bytes(&pages.to_bytes()?)?;
let renamed = focused
    .edit_body_table_name(selector)?
    .set_name("Revenue by Quarter")?
    .commit()?;
let mut renamed_bytes = Vec::new();
renamed.package().write_to(&mut renamed_bytes)?;
pages = PagesEditor::from_bytes(&renamed_bytes)?;
pages.resize_table(table.model_object_id, 5, 4)?;
let second_anchor = pages.body_text()?.encode_utf16().count();
let notes = pages.add_table(second_anchor, "Notes", 2, 2)?;
pages.set_table_cell(
    notes.model_object_id,
    0,
    0,
    CellValue::Text("Generated independently".to_owned()),
)?;
pages.save("created-with-table.pages")?;
# Ok::<(), Box<dyn std::error::Error>>(())
```

`PagesEditor::add_table` bootstraps the first native table in a scratch-created
document and reuses an existing native style template for later tables. Every
table receives independent cell stores, row and column identities, and formula
ownership. The same table operations work on app-created files; see
`add_pages_table` and `edit_pages_table`.
`PagesEditor::remove_table` transactionally removes the body anchor and the
table's private storage and formula graph while preserving shared objects and
other tables. See `remove_pages_table` for a complete file-to-file example and
`inspect_pages_tables` for model identifiers and dimensions.

Scratch-created documents can also add independent, body-anchored text boxes;
no existing drawable or template package is required:

```rust
use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

let mut pages = PagesEditor::create_with_text("Quarterly report")?;
pages.add_text_box(
    "Quarterly report".encode_utf16().count(),
    "Prepared from typed IWA objects",
    DrawablePoint { x: 96.0, y: 144.0 },
    DrawableSize { width: 240.0, height: 72.0 },
)?;
pages.save("created-with-text-box.pages")?;
# Ok::<(), litchi_iwa::Error>(())
```

Ordinary text-bearing shapes have independent CRUD. Rectangle, rounded
rectangle, ellipse, left-arrow, right-arrow, double-arrow, regular-polygon, and
star paths are constructed from typed, validated presets together with their
storage, stand-ins, body attachment, z-order, style relationship, and UUIDs. No
source drawable or package is copied:

```rust
use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, Preset};

let body = "Quarterly report";
let mut pages = PagesEditor::create_with_text(body)?;
let shape = pages.add_body_shape(
    body.encode_utf16().count(),
    "A fully editable shape",
    DrawablePoint { x: 180.0, y: 240.0 },
    DrawableSize { width: 300.0, height: 150.0 },
    Preset::RightArrow,
)?;
pages.set_body_shape_text(shape.drawable_object_id, "Updated")?;
pages.set_body_shape_preset(shape.drawable_object_id, Preset::DoubleArrow)?;
let duplicate = pages.duplicate_body_shape(
    shape.drawable_object_id,
    pages.body_text()?.encode_utf16().count(),
)?;
pages.set_body_shape_text(duplicate.drawable_object_id, "Independent copy")?;
pages.save("created-with-shape.pages")?;
# Ok::<(), litchi_iwa::Error>(())
```

Straight lines use validated document-space points and typed native endpoint
styles. Their path, empty writable storage, stand-ins, attachment, z-order,
style inheritance, and UUID graph are all source-built:

```rust
use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, Endpoint, Endpoints};

let mut pages = PagesEditor::create_with_text("Built without a template")?;
let line = pages.add_body_line_with_endpoints(
    pages.body_text()?.encode_utf16().count(),
    DrawablePoint { x: 180.0, y: 240.0 },
    DrawablePoint { x: 480.0, y: 390.0 },
    Endpoints::new(Endpoint::OpenCircle, Endpoint::FilledArrow),
)?;
pages.set_body_line_segment(
    line.drawable_object_id,
    DrawablePoint { x: 96.0, y: 180.0 },
    DrawablePoint { x: 456.0, y: 180.0 },
)?;
assert_eq!(
    pages.body_line_endpoints(line.drawable_object_id)?.end,
    Endpoint::FilledArrow,
);
// pages.reset_body_line_endpoints(line.drawable_object_id)?; // delete decorations
pages.save("created-with-line.pages")?;
# Ok::<(), litchi_iwa::Error>(())
```

Images use the same source-free path. The image object, body attachment,
stand-ins, z-order, style link, UUIDs, component data reference, and `Data/*`
asset are constructed directly; no blank Pages package is embedded:

```rust
use std::fs;
use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

let body = "Quarterly report";
let image = fs::read("chart.png")?;
let mut pages = PagesEditor::create_with_text(body)?;
let source = pages.add_body_image(
    body.encode_utf16().count(),
    "chart.png",
    &image,
    PagesImageOptions::new(
        DrawablePoint { x: 96.0, y: 144.0 },
        DrawableSize { width: 300.0, height: 225.0 },
    ),
)?;
pages.set_body_image_title(source.drawable_object_id, "Quarterly revenue")?;
pages.set_body_image_caption(source.drawable_object_id, "North America, Q4")?;
let duplicate_anchor = pages.body_text()?.encode_utf16().count();
let duplicate = pages.duplicate_body_image(source.drawable_object_id, duplicate_anchor)?;
assert_eq!(duplicate.image_data_identifier, source.image_data_identifier);
pages.save("created-with-image.pages")?;
# Ok::<(), litchi_iwa::Error>(())
```

`duplicate_body_image` retains the native Pages relationship: the duplicated
drawable has independent geometry and body anchoring, but both images share one
embedded asset. Updating either image's bytes therefore updates both.

File-backed movies are also body-anchored and source-built. Their video and
poster assets, playback bounds, drawable graph, stand-ins, body attachment,
z-order, style relationship, UUIDs, and component data references are generated
without an input package:

```rust
use std::fs;
use std::time::Duration;
use litchi_iwa::pages::PagesEditor;
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_pages::movie::Options as PagesMovieOptions;

let body = "Quarterly report";
let movie = fs::read("demo.mov")?;
let poster = fs::read("demo-poster.png")?;
let mut pages = PagesEditor::create_with_text(body)?;
let source = pages.add_body_movie(
    body.encode_utf16().count(),
    "demo.mov",
    &movie,
    "demo-poster.png",
    &poster,
    PagesMovieOptions::new(
        Point { x: 96.0, y: 144.0 },
        Size { width: 320.0, height: 180.0 },
        Duration::from_secs(8),
    )?,
)?;
let duplicate_anchor = pages.body_text()?.encode_utf16().count();
let duplicate = pages.duplicate_body_movie(source.drawable_object_id, duplicate_anchor)?;
assert_eq!(duplicate.movie_data_identifier, source.movie_data_identifier);
assert_eq!(
    duplicate.poster_image_data_identifier,
    source.poster_image_data_identifier,
);
pages.save("created-with-movie.pages")?;
# Ok::<(), litchi_iwa::Error>(())
```

`duplicate_body_movie` produces an independently positioned and anchored movie
while keeping its video and poster bytes shared with the source, exactly as
Pages' Duplicate command does.

Audio-only media controls use the same source-free body ownership model. The
audio asset, playback bounds, zero-size control geometry, body attachment,
stand-ins, z-order, style relationship, UUIDs, and component data reference are
created directly from typed objects:

```rust
use std::fs;
use std::time::Duration;
use litchi_iwa::pages::PagesEditor;
use litchi_iwa_common::shape::geometry::Point;
use litchi_pages::audio::Options as PagesAudioOptions;

let body = "Interview notes";
let audio = fs::read("interview.aiff")?;
let mut pages = PagesEditor::create_with_text(body)?;
let source = pages.add_body_audio(
    body.encode_utf16().count(),
    "interview.aiff",
    &audio,
    PagesAudioOptions::new(Point { x: 180.0, y: 240.0 }, Duration::from_secs(30))?,
)?;
let duplicate_anchor = pages.body_text()?.encode_utf16().count();
let duplicate = pages.duplicate_body_audio(source.drawable_object_id, duplicate_anchor)?;
assert_eq!(duplicate.audio_data_identifier, source.audio_data_identifier);
pages.save("created-with-audio.pages")?;
# Ok::<(), litchi_iwa::Error>(())
```

`duplicate_body_audio` creates an independently positioned body attachment at
Pages' native 30-point duplicate offset while sharing the audio asset.

### Create Numbers spreadsheets from scratch

Scratch-created spreadsheets can add ordinary text boxes directly to a sheet.
The complete drawable graph and metadata are generated from typed values; no
existing text box or blank package is required:

```rust
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersSheetImageOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

let mut numbers = NumbersDocumentBuilder::new()
    .sheet_name("Forecast")
    .table_name("Revenue")
    .build()?;
let sheet_id = numbers.sheets()?[0].id();
numbers.add_sheet_text_box(
    sheet_id,
    "Prepared from typed IWA objects",
    DrawablePoint { x: 40.0, y: 300.0 },
    DrawableSize { width: 300.0, height: 72.0 },
)?;
numbers.save("created-with-text-box.numbers")?;
# Ok::<(), litchi_iwa::Error>(())
```

Ordinary text-bearing shapes have independent CRUD. Rectangle, rounded
rectangle, ellipse, left-arrow, right-arrow, double-arrow, regular-polygon, and
star paths are constructed from typed, validated presets together with their
storage, stand-ins, style relationship, ownership, and UUIDs. No source
drawable or package is copied:

```rust
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, Preset};

let mut numbers = NumbersDocumentBuilder::new().build()?;
let sheet_id = numbers.sheets()?[0].id();
let shape = numbers.add_sheet_shape(
    sheet_id,
    "A fully editable shape",
    DrawablePoint { x: 420.0, y: 300.0 },
    DrawableSize { width: 300.0, height: 150.0 },
    Preset::RightArrow,
)?;
numbers.set_sheet_shape_text(sheet_id, shape.drawable_object_id, "Updated")?;
numbers.set_sheet_shape_preset(sheet_id, shape.drawable_object_id, Preset::DoubleArrow)?;
let duplicate = numbers.duplicate_sheet_shape(sheet_id, shape.drawable_object_id)?;
numbers.set_sheet_shape_text(sheet_id, duplicate.drawable_object_id, "Independent copy")?;
numbers.save("created-with-shape.numbers")?;
# Ok::<(), litchi_iwa::Error>(())
```

Straight lines are constructed from validated sheet-space points and typed
native endpoint styles. The path, empty storage, stand-ins, ownership, style
inheritance, and UUID graph are emitted without a source package:

```rust
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, Endpoint, Endpoints};

let mut numbers = NumbersDocumentBuilder::new().build()?;
let sheet_id = numbers.sheets()?[0].id();
let line = numbers.add_sheet_line_with_endpoints(
    sheet_id,
    DrawablePoint { x: 420.0, y: 300.0 },
    DrawablePoint { x: 720.0, y: 450.0 },
    Endpoints::new(Endpoint::FilledCircle, Endpoint::SimpleArrow),
)?;
numbers.set_sheet_line_segment(
    sheet_id,
    line.drawable_object_id,
    DrawablePoint { x: 72.0, y: 180.0 },
    DrawablePoint { x: 432.0, y: 180.0 },
)?;
assert_eq!(
    numbers
        .sheet_line_endpoints(sheet_id, line.drawable_object_id)?
        .start,
    Endpoint::FilledCircle,
);
// numbers.reset_sheet_line_endpoints(sheet_id, line.drawable_object_id)?;
numbers.save("created-with-line.numbers")?;
# Ok::<(), litchi_iwa::Error>(())
```

Images are also constructed directly as sheet-owned drawables. The private
image graph, stylesheet link, UUIDs, component data reference, and `Data/*`
asset are generated from typed values; no blank Numbers package is embedded:

```rust
use std::fs;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

let image = fs::read("chart.png")?;
let mut numbers = NumbersDocumentBuilder::new().build()?;
let sheet_id = numbers.sheets()?[0].id();
let source = numbers.add_sheet_image(
    sheet_id,
    "chart.png",
    &image,
    NumbersSheetImageOptions::new(
        DrawablePoint { x: 420.0, y: 180.0 },
        DrawableSize { width: 320.0, height: 240.0 },
    ),
)?;
numbers.set_sheet_image_title(sheet_id, source.drawable_object_id, "Quarterly revenue")?;
numbers.set_sheet_image_caption(sheet_id, source.drawable_object_id, "North America, Q4")?;
let duplicate = numbers.duplicate_sheet_image(sheet_id, source.drawable_object_id)?;
assert_eq!(duplicate.image_data_identifier, source.image_data_identifier);
numbers.save("created-with-image.numbers")?;
# Ok::<(), litchi_iwa::Error>(())
```

`duplicate_sheet_image` follows Numbers' Duplicate command: it creates an
independently positioned drawable while retaining a shared embedded image asset.
Replacing the data through either image updates the pair.

File-backed movies are likewise sheet-owned and source-built. Their video and
poster assets, drawable graph, playback bounds, media style, stand-ins, UUIDs,
and component data references are generated without an input package:

```rust
use std::fs;
use std::time::Duration;
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersSheetMovieOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

let movie = fs::read("demo.mov")?;
let poster = fs::read("demo-poster.png")?;
let mut numbers = NumbersDocumentBuilder::new().build()?;
let sheet_id = numbers.sheets()?[0].id();
let source = numbers.add_sheet_movie(
    sheet_id,
    "demo.mov",
    &movie,
    "demo-poster.png",
    &poster,
    NumbersSheetMovieOptions::new(
        DrawablePoint { x: 420.0, y: 180.0 },
        DrawableSize { width: 320.0, height: 180.0 },
        Duration::from_secs(8),
    ),
)?;
let duplicate = numbers.duplicate_sheet_movie(sheet_id, source.drawable_object_id)?;
assert_eq!(duplicate.movie_data_identifier, source.movie_data_identifier);
assert_eq!(
    duplicate.poster_image_data_identifier,
    source.poster_image_data_identifier,
);
numbers.save("created-with-movie.numbers")?;
# Ok::<(), litchi_iwa::Error>(())
```

`duplicate_sheet_movie` follows Numbers' native 10-point placement and keeps
the duplicate's video and poster assets shared with the original.

Audio-only media controls use the same source-built sheet ownership model. The
audio asset, playback bounds, zero-size control geometry, stand-ins, media
style, UUIDs, and component data reference are emitted directly from typed
objects:

```rust
use std::fs;
use std::time::Duration;
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersSheetAudioOptions};
use litchi_iwa::shapes::DrawablePoint;

let audio = fs::read("interview.aiff")?;
let mut numbers = NumbersDocumentBuilder::new().build()?;
let sheet_id = numbers.sheets()?[0].id();
let source = numbers.add_sheet_audio(
    sheet_id,
    "interview.aiff",
    &audio,
    NumbersSheetAudioOptions::new(
        DrawablePoint { x: 420.0, y: 180.0 },
        Duration::from_secs(30),
    ),
)?;
let duplicate = numbers.duplicate_sheet_audio(sheet_id, source.drawable_object_id)?;
assert_eq!(duplicate.audio_data_identifier, source.audio_data_identifier);
numbers.save("created-with-audio.numbers")?;
# Ok::<(), litchi_iwa::Error>(())
```

`duplicate_sheet_audio` follows Numbers' native 10-point placement while
keeping the duplicate's audio bytes shared with the source.

### Create Keynote presentations from scratch (legacy host scope)

The source-building and graph-editor APIs below remain in the migration host.
For ordinary text in an existing presentation—its slide title, body, or
speaker notes—use the `litchi-keynote::Package` workflow shown above instead
of a `KeynoteEditor` raw-ID operation.

Builder-only work stays in the host, then hands the completed artifact to the
focused package API for a transition edit. The handoff has no native IDs:

```rust,no_run
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_keynote::Package;

let keynote = KeynoteDocumentBuilder::new().title("Draft").build()?;
let package = Package::from_bytes(&keynote.to_bytes()?)?;
# let _ = package;
# Ok::<(), Box<dyn std::error::Error>>(())
```

The builder materializes Keynote's modern storage-less slide-number placeholder
graph in both the theme layout and live slide. It can be initially visible or
retained hidden for later toggling; fresh slides cloned from the layout preserve
the same native behavior:

```rust
use litchi_iwa::keynote::KeynoteDocumentBuilder;

let mut keynote = KeynoteDocumentBuilder::new()
    .title("Native slide numbers")
    .slide_number_visible(true)
    .build()?;
let layout = keynote.default_slide_layout()?;
keynote.add_slide(layout)?;
keynote.save("created-with-slide-numbers.key")?;
# Ok::<(), litchi_iwa::Error>(())
```

See `create_keynote_slide_numbers` for a complete source-free example.

Source-free presentations can also create typed action builds with an editable
custom speed curve. A cubic curve is normalized from `(0, 0)` to `(1, 1)`, and
the same model reads, updates, and removes the native curve payload:

```rust
use litchi_iwa::keynote::{
    KeynoteBuildSettings, KeynoteBuildTimingCurve, KeynoteDocumentBuilder,
    KeynoteMotionPathPoint, KeynoteRotationDirection,
};

let mut keynote = KeynoteDocumentBuilder::new().title("Custom timing").build()?;
let drawable = keynote.slide_drawables(0)?.into_iter().next().ok_or("slide has no drawable")?;
let settings = KeynoteBuildSettings::rotate_action(720.0, KeynoteRotationDirection::Clockwise)
    .with_custom_timing_curve(KeynoteBuildTimingCurve::cubic(
        KeynoteMotionPathPoint::new(0.18, 0.04),
        KeynoteMotionPathPoint::new(0.82, 0.96),
    ))?;
keynote.add_slide_build(0, drawable.object_id, settings)?;
keynote.save("created-with-custom-timing.key")?;
# Ok::<(), Box<dyn std::error::Error>>(())
```

See `create_keynote_custom_timing_curve` for a complete source-free example.

Scratch-created presentations can add ordinary text boxes directly to any
slide. The shape, text storage, stand-ins, ownership, z-order, and metadata are
encoded from typed values; no existing drawable or blank package is required:

```rust
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

let mut keynote = KeynoteDocumentBuilder::new()
    .title("Quarterly review")
    .subtitle("Created entirely by litchi-iwa")
    .build()?;
keynote.add_slide_text_box(
    0,
    "Revenue grew 24% year over year",
    DrawablePoint { x: 144.0, y: 720.0 },
    DrawableSize { width: 1_200.0, height: 120.0 },
)?;
keynote.save("created-with-text-box.key")?;
# Ok::<(), litchi_iwa::Error>(())
```

Ordinary text-bearing shapes have independent CRUD. Rectangle, rounded
rectangle, ellipse, left-arrow, right-arrow, double-arrow, regular-polygon, and
star paths are constructed from typed, validated presets together with their
storage, stand-ins, style relationship, ownership, z-order, and UUIDs. No source
drawable or package is copied:

```rust
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, Preset};

let mut keynote = KeynoteDocumentBuilder::new().build()?;
let shape = keynote.add_slide_shape(
    0,
    "A fully editable shape",
    DrawablePoint { x: 720.0, y: 660.0 },
    DrawableSize { width: 480.0, height: 240.0 },
    Preset::RightArrow,
)?;
keynote.set_slide_shape_text(0, shape.drawable_object_id, "Updated")?;
keynote.set_slide_shape_preset(0, shape.drawable_object_id, Preset::DoubleArrow)?;
let duplicate = keynote.duplicate_slide_shape(0, shape.drawable_object_id)?;
keynote.set_slide_shape_text(0, duplicate.drawable_object_id, "Independent copy")?;
keynote.save("created-with-shape.key")?;
# Ok::<(), litchi_iwa::Error>(())
```

Straight lines use validated slide-space points, typed endpoint styles, and
the native two-element Bézier representation. Their path, empty storage,
stand-ins, ownership, z-order, style inheritance, and UUIDs are source-built:

```rust
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, Endpoint, Endpoints};

let mut keynote = KeynoteDocumentBuilder::new().build()?;
let line = keynote.add_slide_line_with_endpoints(
    0,
    DrawablePoint { x: 720.0, y: 660.0 },
    DrawablePoint { x: 1_200.0, y: 900.0 },
    Endpoints::new(Endpoint::OpenSquare, Endpoint::FilledDiamond),
)?;
keynote.set_slide_line_segment(
    0,
    line.drawable_object_id,
    DrawablePoint { x: 96.0, y: 108.0 },
    DrawablePoint { x: 456.0, y: 108.0 },
)?;
assert_eq!(
    keynote
        .slide_line_endpoints(0, line.drawable_object_id)?
        .end,
    Endpoint::FilledDiamond,
);
// keynote.reset_slide_line_endpoints(0, line.drawable_object_id)?;
keynote.save("created-with-line.key")?;
# Ok::<(), litchi_iwa::Error>(())
```

Images can be embedded into the same source-free presentation. The image,
stand-ins, ownership, style link, UUIDs, component data reference, and `Data/*`
asset are all created directly; no blank Keynote package is embedded:

```rust
use std::fs;
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_keynote::slide::image::Options as ImageOptions;

let image = fs::read("chart.png")?;
let mut keynote = KeynoteDocumentBuilder::new().build()?;
let source = keynote.add_slide_image(
    0,
    "chart.png",
    &image,
    ImageOptions::new(
        Point { x: 704.0, y: 284.0 },
        Size { width: 512.0, height: 512.0 },
    )?,
)?;
keynote.set_slide_image_title(0, source.drawable_object_id, "Quarterly revenue")?;
keynote.set_slide_image_caption(0, source.drawable_object_id, "North America, Q4")?;
let duplicate = keynote.duplicate_slide_image(0, source.drawable_object_id)?;
assert_eq!(duplicate.image_data_identifier, source.image_data_identifier);
keynote.save("created-with-image.key")?;
# Ok::<(), litchi_iwa::Error>(())
```

`duplicate_slide_image` mirrors Keynote: the duplicate is independently
positioned on the same slide, while its media data stays shared with the source.
Replacing either image's bytes updates both images.

File-backed movies use the same source-built path. The video, poster, media
style, stand-ins, component registrations, and Keynote's automatic playback
build and timing chunk are generated from typed values:

```rust
use std::fs;
use std::time::Duration;
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_keynote::slide::movie::Options as SlideMovieOptions;

let movie = fs::read("demo.mov")?;
let poster = fs::read("demo-poster.png")?;
let mut keynote = KeynoteDocumentBuilder::new().build()?;
let source = keynote.add_slide_movie(
    0,
    "demo.mov",
    &movie,
    "demo-poster.png",
    &poster,
    SlideMovieOptions::new(
        Point { x: 640.0, y: 360.0 },
        Size { width: 640.0, height: 360.0 },
        Duration::from_secs(8),
    )?,
)?;
let duplicate = keynote.duplicate_slide_movie(0, source.drawable_object_id)?;
assert_eq!(duplicate.movie_data_identifier, source.movie_data_identifier);
assert_eq!(
    duplicate.poster_image_data_identifier,
    source.poster_image_data_identifier,
);
keynote.save("created-with-movie.key")?;
# Ok::<(), litchi_iwa::Error>(())
```

`duplicate_slide_movie` mirrors Keynote's Duplicate command: it creates a new
movie graph and playback build with shared video and poster data.

Independently positioned audio uses a distinct typed API even though Keynote
stores it in the movie-archive family. Its zero-size control, media style,
stand-ins, component registrations, and native Start Audio build are created
without copying a package or drawable:

```rust
use std::fs;
use std::time::Duration;
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa_common::shape::geometry::Point;
use litchi_keynote::slide::audio::Options as SlideAudioOptions;

let audio = fs::read("narration.aiff")?;
let mut keynote = KeynoteDocumentBuilder::new().build()?;
let source = keynote.add_slide_audio(
    0,
    "narration.aiff",
    &audio,
    SlideAudioOptions::new(Point { x: 960.0, y: 540.0 }, Duration::from_secs(12))?,
)?;
let duplicate = keynote.duplicate_slide_audio(0, source.drawable_object_id)?;
assert_eq!(duplicate.audio_data_identifier, source.audio_data_identifier);
keynote.save("created-with-audio.key")?;
# Ok::<(), litchi_iwa::Error>(())
```

`duplicate_slide_audio` creates a separate audio graph and Start Audio build
while sharing the embedded audio asset, mirroring Keynote's Duplicate command.

All source-built media exposes its shared `DrawableProperties` without
normalizing unrelated native fields. Read the current value, update only the
field you need, then write it back—for example,
`body_movie_properties`/`set_body_movie_properties`,
`sheet_audio_properties`/`set_sheet_audio_properties`, or
`slide_movie_properties`/`set_slide_movie_properties`. The paired Pages,
Numbers, and Keynote movie/audio APIs preserve unknown movie-archive fields and
carry the properties through native-style duplication.

The same media APIs expose the archive-free
`litchi_iwa_common::media::playback::{MediaPlaybackSettings, MediaVolume,
MediaLoopMode}` vocabulary for typed trim boundaries, poster position, repeat
mode, and volume. Body and sheet media continue to use their matching host
`*_playback_settings` methods. Keynote slide-media playback is owned by
`litchi_keynote::Package::edit_slide_movie_playback_settings` and its typed
`SlideSelector`/`MovieSelector` selectors; audio controls are included in that
source-ordered media collection. The update preserves unrelated and unknown
movie-archive fields; the common builders reject invalid levels and trim
ranges, and `MediaLoopMode::Unknown` allows a newer native repeat value to
round-trip.

### Edit existing documents through the migration host

This compatibility example uses host APIs only for operations that have not
yet moved to a concrete package owner. It does not demonstrate Keynote title,
body, or speaker-notes editing: those semantic operations are owned by
`litchi-keynote`.

#### Keynote slide-table discovery boundary

Legacy `KeynoteEditor::slide_tables` now uses a private bounded catalog for
its package-wide discovery pass, and `add_slide_table` reuses the same bounded
template-discovery route. The catalog keeps only compact object/member slots
and message type/length facts; decompressed archives and payloads are borrowed
for strict projections and are not retained. A borrowed model projection
supplies only the table identity, name, and dimensions needed by discovery,
with finite archive/object/message/reference/payload/retained/semantic and
wire limits. Canonical model type 6001 is authoritative; strict legacy
type-6000 is considered only when no type-6001 model exists, and simultaneous
model candidates fail closed.

This is a discovery optimization and admission boundary, not a new semantic
Keynote package API. The legacy `KeynoteSlideTableInfo` result continues to
expose its compatibility native identifiers. Complete TableInfo/model
materialization, appearance and lock handling, storage/cells, formulas, tiles,
comments, metadata, and native mutations remain on their existing generated
or native compatibility paths. The private catalog does not claim a fully
generated-free Keynote implementation, a zero-copy package, or retirement of
the migration host, its dependency edges, or migration debts.

An external public `KeynoteEditor` read-only driver checked the Wave99 source
`/private/tmp/wave99-native-table-source.key` (499,854 bytes, SHA-256
`e5ddda5583d4312f67501312f2681859e0969bc8c49cdcd8160f9b93fe4c21ef`) and the
Wave100 source (500,128 bytes, SHA-256
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`). Each
read one slide and one 5-by-4 `Table 1`; simultaneous 6000/6001 candidates
reject; fresh reopen parity was true and
post-read bytes and hashes were unchanged. Computer Use also opened the
Wave100 source in Keynote 14.4 without repair, recovery, or conversion UI and
observed `Table 1` with 5 rows and 4 columns; source bytes and hash remained
exact. This is bounded read-only semantic and open/render evidence only: no
Keynote save/normalization/mutation, performance/RSS, or public
catalog-statistics claim follows.

#### Keynote slide-table appearance listing boundary

The same bounded catalog-backed listing path now projects table appearance
through strict borrowed `table_appearance_codec` views. It follows the model
style/preset, preset-to-network, network-to-style, and bounded full
parent-inheritance routes; a direct nonzero style retains legacy precedence.
Missing, malformed, duplicate, role-aliased, cyclic, and over-depth facts on
the projected model/style/preset/network and traversed parent routes fail
closed. Stylesheet-registry ownership and unprojected style-property fields
remain compatibility-owned. Appearance is listing support only: existing
public APIs and writers, legacy mutation, and generated TableInfo geometry
remain at their current owners. No metadata, UUID, save-token,
global-ownership, or full-package budget behavior is added, and the
catalog/projection does not claim zero-copy, allocation-free operation, RSS,
or wall-clock performance.

The focused gates are 14/14 for the appearance codec, 50/50 for the
Keynote slide-table suite, 598/598 for the boundary suite, and an empty live
audit; protos/IWA checks, strict scoped Clippy (with existing Pages warnings
allowed for IWA), isolated appearance fuzz checking, Rust 2024 formatting,
and diff checks passed.

An external read-only `KeynoteEditor` driver replayed the Wave99 and Wave100
sources and found one 5-by-4 `Table 1` in each. Both reported row banding
disabled, fixed row sizing, and all five gridline classes visible; fresh
replay was exact across all 86/86 ZIP members, with no changed, added, or
removed member. Computer Use opened and rendered an immutable disposable
Wave100 copy in Keynote 14.4 without repair, recovery, or conversion UI; the
selected table showed alternating row color off, resize rows to fit off, all
five gridline toggles on, 1 header column, 2 header rows, 1 footer row, and
5-by-4 dimensions. Its 500,128-byte source hash remained exact after close.
A separate writable disposable copy auto-persisted view state on close, so
the byte-identity statement applies only to the immutable copy. No native
appearance mutation, save, normalization, or post-save acceptance is claimed.

#### Keynote slide-table appearance owner

Selector-first `litchi-keynote` APIs now read and edit admitted slide-table
appearance through `Package::slide_table_appearance`,
`edit_slide_table_appearance`, and `apply_slide_table_appearance`. A direct
nonzero style uses same-component copy-on-write with prepared appearance
codecs and strict catalog/metadata/UUID/external-edge/ArchiveInfo checks,
candidate reopen, exact inverse, and locality verification. Preset/network
routes remain read-only; preset-only mutation fails closed.

Missing, malformed, duplicate, role-aliased, wrong-wire, cyclic, over-depth,
ambiguous, or otherwise unproven canonical routes fail closed. Existing
generated TableInfo geometry, public compatibility APIs, and legacy/native
mutation remain at their current owners. This boundary does not mutate
physical rows, cells, formulas, storage, tiles, or unrelated table graphs;
changed commits delete stale previews instead of rewriting them. Its resource
ledger is operation-local logical accounting, not
allocator, cache, RSS, or package-wide budget telemetry; it makes no claim of
full generated-free ownership, zero-copy, allocation-free behavior, or
wall-clock performance.

ArchiveInfo admission resolves every distinct referenced object and current
data identifier while preserving native duplicate occurrences in unrelated
producer FieldInfo lists; the selected slide, model, and style routes remain
exact-one checks.

The focused appearance integration passed 19/19; the Keynote library check
and strict library/test Clippy passed. Host bridge, fuzz, and boundary checks
remain within the scoped coverage; this is not a full-workspace-green claim.

The package driver used
`/private/tmp/wave100-native-table-headers-source.key` (500,128 bytes,
SHA-256
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`) and
produced a 466,448-byte candidate (SHA-256
`eb51188fefe44f13001bb2db24fdc4fde4f4bdc208dbc98f6e8b2e4a4a0b1b28`), an
exact inverse, and an exact no-op. The candidate reopened semantically as
banding enabled, fit-cell-content row sizing, and all gridlines hidden;
diagnostics were changed=true, touched_components=3, deleted_previews=3,
and full_reparse=true, with 83 members after deleting three previews.

Computer Use opened a disposable changed-candidate copy in Keynote without
repair, recovery, or conversion UI and showed alternating rows on,
resize-to-fit on, and all five gridline checkboxes off. Keynote normalized
only that disposable copy on close, changing the 466,448-byte copy from the
candidate SHA-256 to 496,491 bytes with SHA-256
`280c57e3e997ef2380edd7759fc2ff17f7df9ee09fbdcbec2978b4b4ce751f57`.
The canonical source, candidate, and inverse remained untouched. No native
mutation, save, reopen acceptance, or byte-exact native UI-save claim follows.

Focused Numbers scalar-cell operations are selector-first `litchi-numbers`
package transactions, not `NumbersEditor` raw-ID convenience calls. Generic
source-built or cross-format `DataFormat` compatibility remains a
migration-host surface. Focused transactions stage a complete batch before
publication, so any rejected coordinate, dependency, or cache update leaves
the package unchanged:

Existing-cell Number, Percentage, Currency, Scientific, Fraction, Text, Date
& Time, and Duration format transactions follow the same focused-package rule.
The former `NumbersEditor` raw-ID convenience methods are retired for Number,
Percentage, Currency, Scientific, Fraction, Date & Time, and Duration; their
dedicated `NumbersEditor` Number/Percentage/Currency/Scientific/Fraction bridge and fallback
entry points are retired as well; private attached Pages/Keynote adapters remain. There is no
dedicated host raw-ID Text
route; the focused Text owner handles that existing-cell operation. Generic
source-built or cross-format `DataFormat` compatibility helpers, including
`DataFormat::Duration`, source-built `DataFormat::Text` and
`DataFormat::DateTime`, the broad `TextDateTimeField` smart-field lifecycle,
and attached Pages/Keynote table callers remain
migration-host-only compatibility surfaces. The dedicated `NumbersEditor`
semantic/table Date & Time and Duration get/set/reset methods are retired;
private Date & Time and Duration bridge helpers remain only behind attached
PagesEditor and KeynoteEditor table compatibility wrappers.

The focused `litchi-numbers` Text-format owner handles existing-cell Text
through its selector-first package API. For source-built Numbers compatibility,
the generic `DataFormat::Text` path remains host-owned, while the attached Pages
and Keynote table callers remain available for their still-hosted cross-format
compatibility surfaces.

The focused `litchi-numbers` Date & Time owner likewise handles one existing
cell through selector-first `Package` APIs, with a bounded native date/time pattern string
and strict native type-261 admission. Its byte bound and field
envelope are checked, but pattern grammar is not validated. This does not replace the broad host
`TextDateTimeField` smart-field lifecycle described below: its raw/native
identity-bearing path remains available for compatibility, as do the Pages and
Keynote wrappers. Native evidence for the focused owner is limited to the one
recorded Numbers type-9 cell/pattern and its operation-specific save/close/
reopen cycle; it is not a suite-wide DateTime or host-exit claim.

The focused `litchi-numbers` Duration owner likewise handles one existing cell
through selector-first `Package` APIs, with strict native type-268 admission
and explicit marker `0x0004`/`0x0005` shapes. Only a true
marker/kind/reference-free absence reads as `None`; marker-zero tuples that
retain Duration kind/reference metadata are ambiguous inherited state and fail
closed. An automated AppleScript-driven Numbers 14.4 open/save/close/reopen
probe successfully round-tripped both admitted marker forms with no reported
error or repair/conversion indication; no GUI repair-dialog inspection was
performed. Strict semantic no-op rereads, exact inverse restoration, and recorded
candidate/native-resaved and Duration Tile/DataList member hashes matched (see
[ADR 0008](../../docs/adr/0008-migration-and-verification.md#2026-09-04-amendment-numbers-existing-cell-duration-owner-and-raw-id-host-route-retirement)).
This is disposable, operation-specific E3/E4 evidence only: the Apple-authored
source/probe artifacts are provenance, not a checked-in E2 fixture, and the
record does not claim broad native acceptance, package parity, or a suite-wide
host-exit result. Generic source-built/cross-format `DataFormat::Duration` and
the attached Pages/Keynote table wrappers remain host-owned compatibility
surfaces.

The focused `litchi-numbers` Custom owner handles one existing rooted cell's
document-scoped Number, Text, or Date & Time custom format through selector-first
`Package` transactions. Its admitted registry routes are
`TN.DocumentArchive` field 9/message type 222 and `TN.super` field 8 →
`TSA.custom_format_list` field 12. Native IDs, registry UUIDs, format-list
keys, archive members, and wire payloads remain private. Deterministic
source-built exact-source fixtures and strict codec/package tests are joined by
operation-specific Numbers 14.4 replacement/clear evidence; focused validation
passes 37 integration tests and boundary verification passes 927 policy tests.
The three dedicated raw-ID `NumbersEditor` Custom conveniences are retired.
Generic source-built/cross-format `DataFormat::Custom` compatibility and
private attached Pages/Keynote table adapters remain host-owned, and a
focused-owner refusal is terminal rather than a fallback trigger. The native
resaved artifacts are recorded in the [fixture README](../../test-data/iwork/README.md#numbers-custom-format-raw-id-retirement-2026-09-06).

```rust,no_run
use litchi_numbers::{Package, SheetSelector, TableSelector};
use litchi_numbers::table::cells::Input;

let package = Package::open("input.numbers")?;
let commit = package
    .edit_table_cells(
        SheetSelector::name("Summary"),
        TableSelector::name("Revenue"),
    )?
    .set_a1("B3", Input::number(42.0)?)?
    .set_a1("C3", Input::text("Revised")?)?
    .clear_a1("D3")?
    .commit()?;

let restored = commit
    .package()
    .apply_table_cells(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

See `litchi-numbers/examples/edit_table_cells.rs` for bounded batch parsing
and synchronized sibling-temporary, distinct-output, no-clobber publication.
Selector-first `litchi_numbers::Package` reads provide ID-free root-comment
and admitted direct-reply projections. The compatibility example below
intentionally retains deprecated native-ID `NumbersEditor` calls for reply
creation, replacement, removal, and graph cleanup.

```rust
#![allow(deprecated)]

use litchi_iwa::numbers::{
    FormulaAxisReference, FormulaCellReference, FormulaExpression, NumbersEditor,
};
use litchi_iwa::pages::PagesEditor;
use litchi_numbers::{SheetSelector, TableSelector};
use litchi_pages::header_footer::Kind;
use litchi_pages::Package as PagesPackage;
use litchi_iwa::keynote::{
    KeynoteBuildSettings, KeynoteBuildStart, KeynoteEditor, KeynoteFlipDirection,
    KeynoteHorizontalBuildDirection, KeynoteKeyboardDirection, KeynoteRotationDirection,
    KeynoteSlideTextRole, KeynoteSwooshDirection,
};

let mut numbers = NumbersEditor::open("input.numbers")?;
let table = numbers.tables()?.remove(0);
numbers.set_cell_comment(table.id(), 1, 3, "Check this value")?;
let _comment = numbers.cell_comment(table.id(), 1, 3)?;
let reply_id = numbers.add_cell_comment_reply(table.id(), 1, 3, "Looks good")?;
let reply_id = numbers.set_cell_comment_reply(
    table.id(),
    1,
    3,
    reply_id,
    "Verified",
)?;
numbers.remove_cell_comment_reply(table.id(), 1, 3, reply_id)?;
numbers.clear_cell_comment(table.id(), 1, 3)?;
numbers.set_formula(
    table.id(),
    2,
    2,
    FormulaExpression::function(
        "SUM",
        [FormulaExpression::Number(1.0), FormulaExpression::Number(2.0)],
    ),
)?;
numbers.set_formula(
    table.id(),
    4,
    2,
    FormulaExpression::function(
        "SUM",
        [FormulaExpression::columns(
            FormulaAxisReference::relative(0),
            FormulaAxisReference::absolute(1),
        )],
    ),
)?;
numbers.set_formula(
    table.id(),
    3,
    2,
    FormulaExpression::function(
        "SUM",
        [FormulaExpression::range(
            FormulaCellReference::relative(0, 0),
            FormulaCellReference::absolute(1, 1),
        )],
    ),
)?;
let pivot_categories = numbers.pivot_categories()?;
// Each entry's typed `reference` can be passed to
// `FormulaExpression::pivot_category` when editing a pivot value formula.
assert!(pivot_categories.iter().all(|category| category.label.is_some()));
numbers.resize_table(table.id(), 30, 10)?;
let original_sheet = numbers.sheets()?[0].clone();
let copied_sheet = numbers.duplicate_sheet(SheetSelector::index(original_sheet.index))?;
numbers.remove_sheet(SheetSelector::index(copied_sheet.index))?;
let new_sheet = numbers.add_empty_sheet("Archive")?;
let new_table = numbers.add_empty_table(SheetSelector::index(new_sheet.index), "Log", 100, 6)?;
numbers.remove_table(TableSelector::name(&new_table.name))?;
numbers.remove_sheet(SheetSelector::index(new_sheet.index))?;
if let Some(sheet) = numbers.sheets()?.first()
    && let Some(text_box) = numbers.sheet_text_boxes(sheet.id())?.first()
{
    numbers.set_sheet_text_box_text(
        sheet.id(),
        text_box.drawable_object_id,
        "Updated text box",
    )?;
    let geometry =
        numbers.sheet_text_box_geometry(sheet.id(), text_box.drawable_object_id)?;
    numbers.set_sheet_text_box_geometry(
        sheet.id(),
        text_box.drawable_object_id,
        geometry,
    )?;
    let properties =
        numbers.sheet_text_box_properties(sheet.id(), text_box.drawable_object_id)?;
    numbers.set_sheet_text_box_properties(
        sheet.id(),
        text_box.drawable_object_id,
        properties,
    )?;
    numbers.set_sheet_drawable_comment(
        sheet.id(),
        text_box.drawable_object_id,
        "Review this text box",
    )?;
    let _comment =
        numbers.sheet_drawable_comment(sheet.id(), text_box.drawable_object_id)?;
    numbers.clear_sheet_drawable_comment(sheet.id(), text_box.drawable_object_id)?;
    let copy = numbers.duplicate_sheet_text_box(
        sheet.id(),
        text_box.drawable_object_id,
        "Independent copy",
    )?;
    numbers.remove_sheet_text_box(sheet.id(), copy.drawable_object_id)?;
}
numbers.save("updated.numbers")?;

let pages_package = PagesPackage::open("input.pages")?;
let first_header = (PagesPackage::header_footers)(&pages_package)?
    .into_iter()
    .find(|region| region.kind() == Kind::Header)
    .expect("document header");
let mut header_edit = pages_package.edit_header_footer_text(first_header.selector())?;
header_edit.set("Quarterly report")?;
let header_commit = header_edit.commit()?;
let mut updated_pages = Vec::new();
header_commit.package().write_to(&mut updated_pages)?;
let mut pages = PagesEditor::from_bytes(&updated_pages)?;
let section_id = pages.sections()[0].object_id;
// Section-name editing now lives in litchi-pages and uses SectionSelector;
// see litchi-pages/examples/edit_section_name.rs.
// Aggregate section settings, including header/footer inheritance and the
// first-page flags, now live in litchi-pages; see
// litchi-pages/examples/edit_section_settings.rs.
// Section-pagination editing now lives in litchi-pages and uses
// SectionSelector; see litchi-pages/examples/edit_section_pagination.rs.
// Section-background editing now lives in litchi-pages and uses
// SectionSelector; see litchi-pages/examples/edit_section_background.rs.
let inserted = pages.insert_section(section_id, 8, "Methods")?;
pages.remove_section(inserted.object_id)?;
let appended = pages.append_section(section_id, "Appendix")?;
pages.remove_section(appended.object_id)?;
if let Some(text_box) = pages.drawable_text_storages()?.first() {
    pages.set_drawable_text(text_box.drawable_object_id, "Updated text box")?;
    let geometry = pages.text_box_geometry(text_box.drawable_object_id)?;
    pages.set_text_box_geometry(text_box.drawable_object_id, geometry)?;
    let properties = pages.text_box_properties(text_box.drawable_object_id)?;
    pages.set_text_box_properties(text_box.drawable_object_id, properties)?;
    let copy = pages.duplicate_text_box(text_box.drawable_object_id, 0, "Independent copy")?;
    pages.remove_text_box(copy.drawable_object_id)?;
}
if let Some(drawable) = pages.drawables()?.first() {
    pages.set_drawable_comment(drawable.object_id, "Review this object")?;
    let _comment = pages.drawable_comment(drawable.object_id)?;
}
pages.save("updated.pages")?;

let mut keynote = KeynoteEditor::open("input.key")?;
if let Some(text_box) = keynote
    .slide_text_storages(0)?
    .into_iter()
    .find(|text| text.role == KeynoteSlideTextRole::TextBox)
{
    keynote.set_slide_text_storage(0, text_box.drawable_object_id, "Updated text box")?;
    let geometry = keynote.slide_text_box_geometry(0, text_box.drawable_object_id)?;
    keynote.set_slide_text_box_geometry(0, text_box.drawable_object_id, geometry)?;
    let properties = keynote.slide_text_box_properties(0, text_box.drawable_object_id)?;
    keynote.set_slide_text_box_properties(0, text_box.drawable_object_id, properties)?;
    let copy = keynote.duplicate_slide_text_box(
        0,
        text_box.drawable_object_id,
        "Independent copy",
    )?;
    keynote.remove_slide_text_box(0, copy.drawable_object_id)?;
}
let layout = keynote.default_slide_layout()?;
keynote.add_slide(layout)?;
if let Some(drawable) = keynote.slide_drawables(0)?.first() {
    keynote.set_slide_drawable_comment(0, drawable.object_id, "Review this slide object")?;
    let _comment = keynote.slide_drawable_comment(0, drawable.object_id)?;

    let build = keynote.add_slide_build(
        0,
        drawable.object_id,
        KeynoteBuildSettings::appear_in(),
    )?;
    let mut build_settings = KeynoteBuildSettings::dissolve_in();
    build_settings.duration = 1.5;
    build_settings.start = KeynoteBuildStart::AfterTransition;
    build_settings.delay = 0.25;
    keynote.set_slide_build(0, build.object_id, build_settings)?;
    let _builds = keynote.slide_builds(0)?;
    keynote.remove_slide_build(0, build.object_id)?;

    let build_out = keynote.add_slide_build(
        0,
        drawable.object_id,
        KeynoteBuildSettings::appear_out(),
    )?;
    keynote.remove_slide_build(0, build_out.object_id)?;

    let rotate = keynote.add_slide_build(
        0,
        drawable.object_id,
        KeynoteBuildSettings::rotate_action(810.0, KeynoteRotationDirection::Clockwise),
    )?;
    keynote.remove_slide_build(0, rotate.object_id)?;

    let scale = keynote.add_slide_build(
        0,
        drawable.object_id,
        KeynoteBuildSettings::scale_action(1.5),
    )?;
    keynote.remove_slide_build(0, scale.object_id)?;

    let opacity = keynote.add_slide_build(
        0,
        drawable.object_id,
        KeynoteBuildSettings::opacity_action(37.0),
    )?;
    keynote.remove_slide_build(0, opacity.object_id)?;

    let mut move_build = KeynoteBuildSettings::move_action(488.5, -258.2);
    move_build.move_action.as_mut().unwrap().align_to_path = true;
    let move_build = keynote.add_slide_build(0, drawable.object_id, move_build)?;
    keynote.remove_slide_build(0, move_build.object_id)?;

    let pulse = keynote.add_slide_build(
        0,
        drawable.object_id,
        KeynoteBuildSettings::pulse_action(6, 135.0),
    )?;
    keynote.remove_slide_build(0, pulse.object_id)?;

    let flip = keynote.add_slide_build(
        0,
        drawable.object_id,
        KeynoteBuildSettings::flip_action(4, KeynoteFlipDirection::RightToLeft),
    )?;
    keynote.remove_slide_build(0, flip.object_id)?;

    let keyboard = keynote.add_slide_build(
        0,
        drawable.object_id,
        KeynoteBuildSettings::keyboard_in(KeynoteKeyboardDirection::Forward, true),
    )?;
    keynote.remove_slide_build(0, keyboard.object_id)?;

    let trace = keynote.add_slide_build(
        0,
        drawable.object_id,
        KeynoteBuildSettings::trace_in(KeynoteHorizontalBuildDirection::LeftToRight),
    )?;
    keynote.set_slide_build(
        0,
        trace.object_id,
        KeynoteBuildSettings::swoosh_out(KeynoteSwooshDirection::FromRight),
    )?;
    keynote.remove_slide_build(0, trace.object_id)?;
}
keynote.save("updated.key")?;
# Ok::<(), litchi_iwa::Error>(())
```

### Pages document and footnote settings use one focused transaction

Document visibility, facing-page, hyphenation, and ligature options share one
focused transaction with footnote formatter settings. Use the semantic
`litchi_pages::document_settings::Settings` value; it exposes neither native
identifiers nor raw records. Packages and commits are immutable, so carry
`commit.package()` into any later transaction.

```rust,no_run
use litchi_pages::{
    Package,
    document_settings::Settings,
    footnote::{self, Format, Kind},
};

let package = Package::open("input.pages")?;
let current = package.document_settings()?;

let mut options = current.options();
options.set_facing_pages(Some(true));
options.set_automatic_hyphenation(Some(true));
options.set_ligatures_enabled(Some(false));

let footnotes = footnote::Settings {
    kind: Some(Kind::Footnotes),
    format: Some(Format::Roman),
    ..current.footnotes()
};
let settings = Settings::new(options, footnotes)?;

let commit = package
    .edit_document_settings()?
    .set(settings)
    .commit()?;
assert_eq!(commit.package().document_settings()?, settings);

let restored = commit
    .package()
    .apply_document_settings(&commit.patch().inverse())?;
let mut original_bytes = Vec::new();
package.write_to(&mut original_bytes)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original_bytes);
# Ok::<(), Box<dyn std::error::Error>>(())
```

Unchanged settings retain the original source allocation. A changed commit
requires an exact flat package: normalized legacy sources can be read but a
changed edit returns `litchi_pages::document_settings::Error::UnsupportedSource`.
Changed commits invalidate the private layout cache, remove stale root previews,
and fully reopen the candidate; the inverse patch restores the exact original
package. See `litchi-pages/examples/edit_document_settings.rs` for the complete
file-to-file publication workflow.

### Pages page layout uses the focused package transaction

Page dimensions, margins, scale, orientation, and vertical body layout are no
longer `PagesEditor` operations. `litchi_pages::Package` exposes only the
validated semantic `Layout`; it does not expose native IDs, components, or raw
records. The package and every commit are immutable, so begin each later edit
from `commit.package()`.

```rust,no_run
use litchi_pages::{
    Package,
    page_layout::Orientation,
};

let package = Package::open("input.pages")?;
let mut edit = package.edit_page_layout()?;
let mut layout = edit.layout();
layout.set_top_margin(Some(54.0))?;
layout.set_orientation(Some(Orientation::Portrait))?;
edit.set_layout(layout)?;
let commit = edit.commit()?;

assert_eq!(commit.package().page_layout()?, layout);

let restored = commit
    .package()
    .apply_page_layout(&commit.patch().inverse())?;
let mut original_bytes = Vec::new();
package.write_to(&mut original_bytes)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original_bytes);
# Ok::<(), Box<dyn std::error::Error>>(())
```

An unchanged layout reuses the exact source allocation and keeps preview and
view-state caches. A changed layout requires an exact flat package; normalized
legacy sources are read-only for this transaction and return
`PageLayoutError::UnsupportedSource` when changed. Changed commits update the
private derived layout state, invalidate and remove stale root previews, and
fully reopen the candidate before publication. The inverse patch restores the
exact original package. See `litchi-pages/examples/edit_page_layout.rs` for a
complete command-line workflow; it requires a distinct new output path and
publishes through a synchronized sibling temporary file with no-clobber
publication.

### Pages section settings use one selector-first aggregate transaction

Section names, header/footer inheritance, first-page and even/odd template
flags, section-start behavior, and page numbering are no longer aggregate
`PagesEditor` raw-ID operations. Select a section by exact semantic name or
checked position and exchange the complete aggregate
`litchi_pages::section::Settings` value. The direct transaction types live under
`litchi_pages::section::settings`; their public paths contain only the resolved
semantic position, never a native object identifier, member name, protobuf
value, or raw record.

```rust,no_run
use litchi_pages::{Package, SectionSelector, section::Settings};

let package = Package::open("input.pages")?;
let selector = SectionSelector::name("Introduction");
let mut settings: Settings = package.section_settings(selector)?;

// Preserve the name and pagination while changing two optional flags.
settings.set_inherit_previous_header_footer(Some(false));
settings.set_first_page_hides_header_footer(Some(true));
let commit = package
    .edit_section_settings(selector)?
    .set(settings)?
    .commit()?;
assert_eq!(
    &commit.package().section_settings(selector)?,
    commit.patch().after(),
);

let restored = commit
    .package()
    .apply_section_settings(&commit.patch().inverse())?;
let mut original_bytes = Vec::new();
package.write_to(&mut original_bytes)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original_bytes);
# Ok::<(), Box<dyn std::error::Error>>(())
```

Each optional Boolean has three lossless states: `None` removes native field
presence, `Some(false)` retains explicit false, and `Some(true)` retains true.
The name likewise distinguishes absence from an explicitly empty string;
pagination keeps absent values and future native enum discriminants without
allowing known values to masquerade as `Unknown`. The consuming `Edit::set`
validates the complete replacement before it can be committed.

An unchanged replacement is exact no-op and retains the source allocation,
derived layout cache, and previews. A changed exact-source transaction proves
the selected section's template dependencies, rewrites only the selected
Section component, preserves ViewState, the rooted layout cache, and every
canonical root preview exactly, and fully reopens and reverse-reads the
candidate. A changed native-supported edit reports one touched component and
zero deleted previews. Its inverse applies only to the exact committed package
and restores the complete original artifact. Changed legacy nested packages
are refused rather than silently normalized.

Matched Pages 14.4 pairs provide a deliberately narrower UI oracle. In the
second section, changing field 17 from explicit false to true made pages 3 and
4 reuse header/footer 1; changing field 19 from explicit false to true showed
header/footer 2 on page 3 and left page 4 blank; changing field 28 from
explicit false to true left page 3 blank and showed header/footer 2 on page 4.
Each pair changed only that scalar in section object 1732889 (type 10011): the
section header, template objects, and all 36 text storages stayed exact. Pages
opened each Rust artifact without a warning and completed Save As, close, and
exact-path reopen with the same observed layout. Field 18 remained explicitly
false in every pair, so first-page-different behavior and absent-versus-false
distinctions are Rust codec/transaction evidence, not native UI claims.
Pages also retained all native previews byte-for-byte, matching the focused
transaction's exact preservation of previews, the layout cache, and every
other ViewState byte.

`edit_section_name` and `edit_section_pagination` remain ergonomic focused
facades when only one value family changes; use the aggregate transaction when
several families must be published atomically. The migration-host
`section_settings`/`set_section_settings` pair and its raw-ID example are
retired. Section backgrounds and section insertion/removal remain separate
migration capabilities. See
`litchi-pages/examples/edit_section_settings.rs` for a selector-first CLI that
preserves name and pagination, verifies exact inverse restoration before
publication, and writes distinct output artifacts through synchronized sibling
temporary files without clobbering existing paths.

### Numbers table locks use the focused package API

Table-lock reads and edits are no longer `NumbersEditor` raw-ID operations.
Use a sheet selector and a table selector scoped to that sheet. The package and
every commit are immutable; start a later edit from `commit.package()`. A patch
is authorized against the exact source package, and applying its inverse to the
committed package restores the original source bytes.

```rust,no_run
use litchi_numbers::table::lock::State as LockState;
use litchi_numbers::{Package, SheetSelector, TableSelector};

let package = Package::open("input.numbers")?;
let sheet = SheetSelector::name("Summary");
let table = TableSelector::name("Revenue");

assert_eq!(package.table_lock(sheet, table)?, LockState::Unlocked);

let mut edit = package.edit_table_lock(sheet, table)?;
edit.lock();
let commit = edit.commit()?;

assert_eq!(
    commit.package().table_lock(sheet, table)?,
    LockState::Locked,
);

let restored = commit
    .package()
    .apply_table_lock(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

### Numbers table headers and footers use the focused package transaction

Header rows and columns, footer rows, frozen headers, and print-time repeated
headers are no longer `NumbersEditor` raw-ID operations. Use the immutable
`Package` with a sheet selector and a table selector scoped to that sheet. The
transaction vocabulary is
`litchi_numbers::table::headers::transaction::{Edit, Patch, Commit, Diagnostics, Error, InvalidReason, LimitKind, Path}`.

```rust,no_run
use litchi_numbers::{
    Package, SheetSelector, TableSelector,
    table::headers::{Count, Settings},
};

let package = Package::open("input.numbers")?;
let sheet = SheetSelector::name("Summary");
let table = TableSelector::name("Revenue");

let mut settings: Settings = package.table_header_settings(sheet, table)?;
settings.header_rows = Some(Count::TWO);
settings.header_columns = Some(Count::ONE);
settings.footer_rows = Some(Count::ONE);
settings.header_rows_frozen = Some(true);
settings.header_columns_frozen = Some(false);
settings.repeating_header_rows_enabled = Some(true);
settings.repeating_header_columns_enabled = Some(false);

let commit = package
    .edit_table_headers(sheet, table)?
    .set(settings)
    .commit()?;

assert_eq!(commit.package().table_header_settings(sheet, table)?, settings);
assert!(commit.diagnostics().changed());

let restored = commit
    .package()
    .apply_table_headers(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

`Settings` preserves field presence: assigning `None` clears an explicitly
stored field to native absence (with an effective count of zero or an effective
boolean of `false`), while `Some(false)` explicitly stores a disabled freeze
or repeat flag. Counts are checked in Numbers' native `1..=5` range, so
`Some(Count::ONE)` represents one footer row. Replacing an edit with its
unchanged `settings()` is an exact no-op; it shares the source package, reports
`changed() == false`, and leaves previews and caches intact. A changed commit
validates the selected table, removes stale root previews, and fully reopens
the candidate; applying the exact-source-checked inverse restores the original
package.

Footer rows and row/column freeze flags are directly supported by this focused
scope. The effective boolean accessors are a native-bool oracle: they report
only the optional stored value (`None` is effectively `false`) and do not infer
interactive state. The transaction deliberately refuses dependent topologies:
any header/footer section-count change with an active pivot or group, a header
row/column count change with a rooted header-name manager, or a repeat-flag
change on a legacy sheet topology returns `Error::UnsupportedDependency`.

Changed edits require an exact flat Numbers package. A legacy nested
`Index.zip` can be read and a no-op preserved, but a changed transaction fails
with `litchi_numbers::table::headers::transaction::Error::UnsupportedSource`
rather than normalizing through the retired migration-host path. See
`litchi-numbers/examples/edit_table_headers.rs` for a distinct-output workflow
that streams with `Package::write_to` through a synchronized sibling temporary
file and publishes without clobbering an existing target.

### Numbers table titles use a focused package transaction

Numbers table-title visibility and outline settings are no longer
`NumbersEditor` raw-ID operations. Use
`litchi_numbers::table::title::{Settings, Edit, Patch, Commit, Diagnostics,
Error, LimitKind, Path}` with a sheet selector and a table selector scoped to
that sheet. `Settings` preserves optional Boolean presence: `None` is absent
on the wire, whereas `Some(false)` is explicitly stored false.
Explicit false and outline presence are losslessly tested transaction values;
they are not native UI-oracle claims.

```rust,no_run
use litchi_numbers::{
    Package, SheetSelector, TableSelector,
    table::title::Settings,
};

let package = Package::open("input.numbers")?;
let sheet = SheetSelector::name("Summary");
let table = TableSelector::name("Revenue");
let before = package.table_title_settings(sheet, table)?;
// This is guaranteed to differ without enabling a previously hidden title.
let settings = if before.visible() == Some(true) {
    Settings::new(None, before.outlined())
} else {
    Settings::new(
        before.visible(),
        match before.outlined() {
            None => Some(false),
            Some(_) => None,
        },
    )
};
let commit = package
    .edit_table_title(sheet, table)?
    .set(settings)
    .commit()?;
assert_eq!(commit.package().table_title_settings(sheet, table)?, settings);
assert!(commit.diagnostics().changed());
assert_eq!(commit.diagnostics().touched_components(), 1);
assert!(commit.diagnostics().deleted_previews() <= 3);

let restored = commit
    .package()
    .apply_table_title(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

Staging the readback value is an exact no-op. A changed title transaction
rewrites the affected `CalculationEngine` component, deletes every existing
canonical root preview, and fully reopens its candidate; the exact-source
inverse restores the original bytes and previews. Changed publication refuses
an effectively locked table, so the changed portion of this example requires
an unlocked supported table. A visible title also requires valid native title
height plus paragraph-style and shape-style prerequisites, so malformed or
unsupported style graphs fail rather than being normalized. The native basic
fixture has all three canonical previews (`3 → 0`, then `0 → 3` on inverse)
and proves only visible `Some(true)` to absent (hide), plus warning-free open
and Save As/reopen; it does not assert outline or explicit-false UI results.
See `litchi-numbers/examples/edit_table_title.rs` for synchronized
sibling-temporary, distinct-output, no-clobber publication through
`Package::write_to`.

Pages body-table titles use the same focused transaction shape through
`litchi_pages::Package::{body_table_title_settings, edit_body_table_title,
apply_body_table_title}` and `BodyTableSelector`; the former raw-object-ID
`PagesEditor` title helpers were removed. Keynote's slide-table title methods
remain migration-host helpers with their format-specific table CRUD.

### Pages body-table appearance uses the focused package transaction

Pages body-table appearance reads and supported exact-source edits use
`litchi_pages::Package::{body_table_appearance,
edit_body_table_appearance, apply_body_table_appearance}` with a typed
`BodyTableSelector` and neutral `table::appearance::Appearance`. Native IDs,
archives, ZIP/wire values, and generated models stay private. The Wave104
owner covers canonical model/style admission, preset/network/default effective
reads, locked exact no-ops, strict ArchiveInfo and opaque-metadata checks,
global style-inbound authority, and source-bound candidate/inverse/locality
verification; unsupported graphs fail closed.

The focused `litchi_pages::Package` owner is strict for all of its reads and
changed writes. Separately, `litchi-iwa` legacy table listing retains one
private read-only appearance helper for compatibility; it does not attempt a
focused Package read, perform mutation, or bypass a focused Package write
failure. Raw `PagesEditor` mutation APIs remain retired, and the appearance
example uses the focused Package owner.

The focused Pages integration passed 21/21, the appearance codec 18/18, the
boundary suite 614/614, and the eight-seed fuzz-target check; strict scoped
Clippy, the `litchi-iwa` library/example checks, and `py_compile` also passed.
The native Pages source was rejected by strict selector read before mutation,
remained byte-exact, and produced no candidate/inverse or UI acceptance
evidence. The operation ledger is logical rather than Package-cache,
decompressed-Archive, allocator, or RSS telemetry. Legacy Pages physical
table/content paths remain in `litchi-iwa`; this note does not imply a
monolith-exit or a full-workspace-green claim.

### Numbers sheet and table names use the focused package transaction

Sheet and table names are no longer `NumbersEditor` raw-ID operations. Use
`litchi_numbers::names::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`
through an immutable `Package`. Both selectors resolve against the original
snapshot, so a sheet and one of its tables can be renamed atomically without
re-resolving the table through the new sheet name.

```rust,no_run
use litchi_numbers::{Package, SheetSelector, TableSelector};

let package = Package::open("input.numbers")?;
let sheet = SheetSelector::name("Summary");
let table = TableSelector::name("Revenue");
let commit = package
    .edit_names()
    .rename_sheet(sheet, "Planning")?
    .rename_table(sheet, table, "Quarterly revenue")?
    .commit()?;

let restored = commit
    .package()
    .apply_names(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

Changed batches preserve selected name-owner records, validate the complete
candidate, and delete the root `preview.jpg`, `preview-micro.jpg`, and
`preview-web.jpg` members when present. A semantic no-op retains the exact
source; read-only legacy nested-`Index.zip` input can likewise remain
preserved, but a changed names transaction returns
`litchi_numbers::names::Error::UnsupportedSource` rather than normalizing it
through the migration host. See `litchi-numbers/examples/edit_names.rs` for a
distinct-output workflow that streams with `Package::write_to` through a
synchronized sibling temporary file and no-clobber publication.

### Numbers sheet order uses a focused package transaction

Moving a sheet is no longer a `NumbersEditor` operation. Use
`litchi_numbers::sheet::order::{Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` with `Package::{edit_sheet_order, apply_sheet_order}`. The selected
sheet is an exact-name or zero-based-position `SheetSelector`, and the
destination is its final zero-based position after removing the selected
sheet. No physical identifiers enter this API.

```rust,no_run
use litchi_numbers::{Package, SheetSelector};

let package = Package::open("input.numbers")?;
let commit = package
    .edit_sheet_order()
    .move_sheet(SheetSelector::name("Archive"), 0)?
    .commit()?;
assert!(commit.diagnostics().changed());
assert_eq!(commit.diagnostics().touched_components(), 1);
assert_eq!(commit.diagnostics().deleted_previews(), 3);
assert!(commit.diagnostics().full_reparse_performed());
assert_eq!(commit.package().sheets()[0].name(), "Archive");

let restored = commit
    .package()
    .apply_sheet_order(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

The transaction atomically keeps both the document order and navigator/sidebar
order in sync. Moving a sheet to its current position is an exact no-op; a
changed move rewrites one owning component, removes the three stale root
previews, and fully reopens its candidate. Its inverse restores all three
previews (`3 → 0` forward, `0 → 3` inverse), accepts only the exact committed
package, and restores the original package bytes. The native gate opens the
changed package warning-free in Numbers, verifies both orders, and confirms a
subsequent Save As/reopen succeeds. See
`litchi-numbers/examples/edit_sheet_order.rs` for synchronized sibling-temp,
distinct-output, no-clobber publication through `Package::write_to`.

Sheet creation, duplication, removal, and moving a table between sheets remain
legacy migration-host capabilities; this focused transaction owns only the
ordering of existing sheets.

Existing Keynote title, body, and speaker-notes storage is owned by the focused
`litchi-keynote` package. Ordinary users must use these APIs rather than the
legacy host's generic storage or raw-ID compatibility paths. They select slides
by semantic position or navigator name, use checked UTF-16 spans, and publish
immutable commits with exact-source-checked inverse patches. Private Buffa
views validate native owner references while bounded raw-record rewriting
preserves untouched bytes.

```rust,no_run
use litchi_keynote::{Package, SlideSelector};

let package = Package::open("input.key")?;
let selector = SlideSelector::index(0);

let mut title = package.edit_slide_title(selector)?;
title.set("Updated title")?;
let title = title.commit()?;

let mut body = title.package().edit_slide_body(selector)?;
body.set("Updated body")?;
let body = body.commit()?;

let mut notes = body.package().edit_slide_notes(selector)?;
notes.set("Presenter cue")?;
let notes = notes.commit()?;
let mut output = Vec::new();
notes.package().write_to(&mut output)?;
assert!(!output.is_empty());
# Ok::<(), Box<dyn std::error::Error>>(())
```

Publish the resulting bytes through the focused
`litchi-keynote/examples/edit_slide_text.rs` example or an equivalent
new-output, sibling-temp, no-clobber operation. The example does not implement
the library's durable atomic-save contract. Never use a truncating
write to replace an existing Keynote package.

### Keynote title and body placeholder visibility is a focused transaction

Per-slide title/body visibility is no longer a `KeynoteEditor` mutation. Use
`litchi_keynote::slide::placeholder::{Kind, State, Edit, Patch, Commit,
Diagnostics, Error, LimitKind}` with a semantic `SlideSelector`; native object
IDs are never exposed. `slide_placeholder_visibility` returns `None` when the
selected title or body placeholder is missing, `Some(State::Visible)` when it
participates in drawing, and `Some(State::Hidden)` when its text and references
remain intact but it is removed from the native drawing and z-order lists.

```rust,no_run
use litchi_keynote::{
    Package, SlideSelector,
    slide::placeholder::{Kind, State},
};

let package = Package::open("input.key")?;
let slide = SlideSelector::name("Opening");
let title_before = package.slide_title(slide)?;
assert_eq!(
    package.slide_placeholder_visibility(slide, Kind::Title)?,
    Some(State::Visible),
);

let commit = package
    .edit_slide_placeholder_visibility(slide, Kind::Title)?
    .set(State::Hidden)
    .commit()?;
assert_eq!(
    commit
        .package()
        .slide_placeholder_visibility(slide, Kind::Title)?,
    Some(State::Hidden),
);
// Hiding changes visibility only; it does not erase the title storage.
assert_eq!(commit.package().slide_title(slide)?, title_before);

let restored = commit
    .package()
    .apply_slide_placeholder_visibility(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

An absent role cannot be edited and returns `Error::PlaceholderNotFound`; this
transaction does not create a placeholder. Its scope is only visibility for
existing title/body layout placeholders: slide-number visibility, layout
editing, and slide or placeholder creation remain separate operations. An
unchanged `set` is an exact no-op; a changed exact-source commit removes stale
root previews and fully reopens the candidate. See
`litchi-keynote/examples/edit_slide_placeholder_visibility.rs` for distinct
output, sibling-temporary, no-clobber publication through `Package::write_to`.

### Keynote per-slide slide-number visibility is a focused transaction

For an existing presentation, per-slide slide-number visibility is no longer a
`KeynoteEditor` mutation. Select the slide semantically and use the same
immutable placeholder transaction with `Kind::SlideNumber`. A read returns
`None` when that slide's layout has no slide-number placeholder,
`Some(State::Visible)` when it draws, and `Some(State::Hidden)` when the
placeholder, its storage/text graph, and its layout reference remain preserved
but it is removed from per-slide drawing and z-order ownership.

```rust,no_run
use litchi_keynote::{
    Package, SlideSelector,
    slide::placeholder::{Kind, State},
};

let package = Package::open("input.key")?;
let slide = SlideSelector::name("Opening");
assert_eq!(
    package.slide_placeholder_visibility(slide, Kind::SlideNumber)?,
    Some(State::Visible),
);

let commit = package
    .edit_slide_placeholder_visibility(slide, Kind::SlideNumber)?
    .set(State::Hidden)
    .commit()?;
assert_eq!(
    commit
        .package()
        .slide_placeholder_visibility(slide, Kind::SlideNumber)?,
    Some(State::Hidden),
);

let restored = commit
    .package()
    .apply_slide_placeholder_visibility(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

This is not the presentation-wide `show::Settings::slide_numbers_visible`
flag: edit that global setting through `Package::{show_settings,
edit_show_settings, apply_show_settings}` when the whole show is in scope. The
visibility transaction neither creates a slide-number placeholder nor changes
its storage/text, layout, or slide creation policy. It is an exact no-op when
the requested state matches; a changed exact-source commit invalidates stale
rendering state and root previews, and its inverse restores the exact source.
See `litchi-keynote/examples/edit_slide_number_visibility.rs` for safe,
distinct-output sibling-temp publication with `Package::write_to` and
no-clobber publication.

### Keynote soundtrack playback is a focused transaction

Playback mode and volume for an existing soundtrack are no longer
`KeynoteEditor` settings mutations. Use
`litchi_keynote::soundtrack::{Mode, Settings, Edit, Patch, Commit,
Diagnostics, Error, LimitKind}` with the immutable
`Package::{soundtrack_settings, edit_soundtrack_settings,
apply_soundtrack_settings}` transaction. A read of `None` means no soundtrack
object exists; `Some(Settings::default())` means it exists with both optional
playback values absent. An edit never creates or deletes a soundtrack object.

```rust,no_run
use std::io;

use litchi_keynote::{
    Package,
    soundtrack::{Mode, Settings},
};

let package = Package::open("input.key")?;
let before = package
    .soundtrack_settings()?
    .ok_or_else(|| io::Error::other("presentation has no soundtrack"))?;
let mut settings: Settings = before;
settings.set_volume(Some(0.35))?;
settings.set_mode(Some(Mode::Loop))?;

let commit = package.edit_soundtrack_settings()?.set(settings).commit()?;
assert!(commit.diagnostics().changed());
assert_eq!(commit.diagnostics().touched_components(), 1);
assert_eq!(commit.package().soundtrack_settings()?, Some(settings));

let restored = commit
    .package()
    .apply_soundtrack_settings(&commit.patch().inverse())?;
assert_eq!(restored.package().soundtrack_settings()?, Some(before));
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

`None` passed to `Settings::set_mode` or `Settings::set_volume` clears native
presence rather than setting a default. `Mode::Unknown(value)` preserves a
genuinely future native mode; known values must use their named variants. An
unchanged commit is an exact no-op. A changed commit rewrites only its owning
soundtrack component, fully reopens the candidate, and leaves rendering
previews intact; its inverse is accepted only by the exact committed package
snapshot. See `litchi-keynote/examples/edit_soundtrack_settings.rs` for
distinct-output, sibling-temporary, synchronized, no-clobber publication
through `Package::write_to`.

This focused transaction owns only playback mode and volume. The migration
host still owns soundtrack-media item and asset operations (including their
ordering and references); use its retained media-item compatibility APIs when
that collection itself must change.

### Keynote existing-slide deletion is a focused transaction

Deleting one existing slide is no longer a `KeynoteEditor` mutation. Use the
immutable `litchi_keynote::slide::delete::{Edit, Patch, Commit, Diagnostics,
Error, LimitKind, Path}` transaction with an exact navigator name or checked
semantic position. The final slide cannot be deleted, native IDs do not enter
the API, and the exact inverse is accepted only by the committed package. The
transaction admits only the supported flat native ownership topology; a
surviving child backlink or other inbound owner is a typed refusal.

```rust,no_run
use litchi_keynote::{Package, SlideSelector};

let package = Package::open("input.key")?;
let mut edit = package.edit_slide_deletion();
edit.remove_slide(SlideSelector::name("Appendix"))?;
let commit = edit.commit()?;
assert_eq!(
    commit.package().show()?.slides().len() + 1,
    package.show()?.slides().len(),
);

let restored = commit
    .package()
    .apply_slide_deletion(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

Deletion removes the selected SlideNode and Slide objects plus their exact
PackageMetadata ownership records. It retains their component registrations,
co-located objects, global data-catalog records, and every media/data payload.
A component data-reference record remains with surviving owners or is removed
when none survive. Shared or newly unreachable media therefore remains
preserved: this operation is structural deletion, not package garbage
collection. See
`litchi-keynote/examples/remove_slide.rs` for distinct-output,
sibling-temporary, synchronized no-clobber publication and optional exact
inverse output.

Keynote slide skip/include, ordering, and modern transition transactions have
focused `litchi-keynote` package-owner paths. Transition transactions use
`litchi_keynote::transition::{Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` and exact-name or typed-position `SlideSelector` values; native
IDs never enter the API. Only an existing modern transition envelope is
editable. `clear` is idempotent and retains Keynote's native no-effect
envelope rather than deleting or synthesizing one. Commits carry an
exact-source-checked inverse patch. A private Buffa lazy view projects known
native fields, while bounded raw-record rewriting preserves accepted source
bytes. The same doc-hidden neutral codec strictly validates retained opaque
color and path payloads for both the focused package and migration adapter.

```rust,no_run
use std::io;

use litchi_keynote::{
    Package, SlideSelector,
    transition::Effect,
};

let package = Package::open("input.key")?;
let selector = SlideSelector::name("Opening");
let mut settings = package
    .slide_transition(selector)?
    .ok_or_else(|| io::Error::other("slide has no modern transition"))?;
settings.set_effect(Some(Effect::Dissolve))?;

let commit = package.edit_slide_transition(selector)?.set(settings)?.commit()?;

let cleared = commit
    .package()
    .edit_slide_transition(selector)?
    .clear()?
    .commit()?;
assert_eq!(
    cleared.package().slide_transition(selector)?.unwrap().effect(),
    Some(&Effect::None),
);
let cleared_again = cleared
    .package()
    .edit_slide_transition(selector)?
    .clear()?
    .commit()?;
assert!(!cleared_again.diagnostics().changed());
let restored = cleared
    .package()
    .apply_slide_transition(&cleared.patch().inverse())?;
assert_eq!(restored.package().slide_transition(selector)?, commit.patch().after().cloned());
let mut output = Vec::new();
cleared_again.package().write_to(&mut output)?;
assert!(!output.is_empty());
# Ok::<(), Box<dyn std::error::Error>>(())
```

See `litchi-keynote/examples/edit_slide_transition.rs` for the focused
distinct-output workflow, which streams the committed package through
`Package::write_to` into a synchronized sibling temporary file before
no-clobber publication.

Text ranges use UTF-16 indexes, matching iWork's attribute tables. Shared
`TSWP.StorageArchive` edits patch only text chunks and affected attribute
indexes. Unknown fields in the storage, attribute tables, entries, ranges, and
references remain byte-exact; removed annotations update reference metadata
without normalizing its ordering. Empty-to-text-to-empty cycles restore real
Pages and Keynote components exactly.

Numbers formulas made from literals, local or cross-table cell, rectangular, whole-row,
whole-column, and pivot-category references, eager built-in functions, unary
operators, and binary operators are compiled to native postfix ASTs and
interned with checked formula-list refcounts. Pivot references are discovered
from the group-by tree and validated against aggregate coordinates, levels,
types, and calculation owners. References preserve whichever inline/tiled
CalculationEngine dependency storage mode the workbook uses. Lazy or volatile
functions, remote data, and spill arrays remain rejected transactionally
because they need additional owner-specific records.
Dependency owners, record tiles, and the engine formula counter are mutated at
bounded wire paths; empty generated tiles are reclaimed, metadata references
are removed, and formula create/clear cycles restore real Numbers archives
byte-for-byte after decompression.

Numeric cells are emitted in Numbers' native decimal128 BNC representation.
Replacing an existing rich-text cell with text preserves its payload identity
and character/paragraph attribute tables when it is unique; shared payloads
are cloned transactionally and rebound only for the target cell. Replacing or
clearing rich text decrements its list reference count and reclaims payload and
storage objects once they are no longer referenced. Segmented string, formula,
and rich-text data lists are read and edited without flattening: existing
entries remain in their segment, segment ranges and reference metadata are
maintained, empty segments are reclaimed, and newly interned entries are added
to the root list with collision-checked identifiers. Formula-error cells
resolve their app-provided error-table text in both legacy and BNC storage;
replacing them clears the cached error identifier and decrements root or
segmented error-list refcounts. Transitional files that
retain complete BNC-v5 mirrors beside legacy pre-BNC rows are validated and
promoted atomically on first edit, making the modern buffers authoritative and
removing the stale legacy row payloads.

Native table cells in Numbers, Pages, and Keynote expose typed horizontal
alignment, validated PostScript font identity, foreground color, underline,
strikethrough, capitalization, normal/superscript/subscript selection, custom
baseline shift, validated character spacing, ligature policy, solid text
background, native outline, drop shadow, native line-spacing modes,
before/after paragraph spacing, first-line/left/right indents, ordered typed
ruler tab stops, and whole-cell point-size, bold, and italic CRUD. These
properties compose in one
paragraph-style variation, use copy-on-write when a style is shared, preserve
unrelated overrides, and reclaim private style objects and list entries when
their last local property is reset. The
`create_iwork_table_layouts` example creates and verifies all three formats
from scratch.

Numbers table-model dimensions are patched at the protobuf wire level.
Unrecognized Apple fields retain their bytes and position, while duplicate
singular fields fail transactionally. Focused header and footer transactions
preserve their native optional presence; see
`litchi-numbers/examples/edit_table_headers.rs`. Full-table sort rules have
typed read, set, and clear APIs that preserve native rule extensions and
reference-tracker metadata.
They configure the order displayed in Numbers' **Organize → Sort** pane;
executing that order remains Numbers' separate **Sort Now** action. Table
resizing still updates tiles,
header buckets, stable UID maps, and stroke sidecars as one checked operation;
each existing object is now mutated through bounded wire paths, including the
unpacked UID index arrays and nested UUID records. Grow/shrink restoration is
byte-exact while unknown fields remain attached to retained rows, headers,
UUIDs, and stroke-layer references.
Physical row and column insertion/deletion also maintains app-authored
`StrokeLayerArchive` border overlays: fixed-axis layers move with their
original cells, crossing runs split or compact around a blank inserted/deleted
axis, and unreachable layer references are removed from the sidecar without
normalizing unknown protobuf fields.
Workbook sheet ordering and standard or form-sheet table ownership lists reuse
the original raw `TSP.Reference` payloads, preserving extensions inside each
reference; newly appended references are removed byte-exactly on rollback or
create/delete cycles.
Focused `litchi_numbers::sheet::order` owns only the order of existing sheets:
it rewrites the document and sidebar order owners together without exposing
their physical identifiers. The migration host retains sheet creation,
duplication, and removal. Public relocation of an existing table belongs to
the focused `litchi_numbers::Package`; the host retains only a crate-private
adapter used while duplicating populated legacy sheets.
`NumbersEditor::add_empty_table` can also recreate the first native table after
the workbook's last table was removed. It derives the style graph from the
workbook theme, builds independent storage and row/column identities, and
registers a fresh formula owner without relying on a hidden template table.
Focused `litchi_numbers::Package::move_table` moves populated tables between
existing sheets without exposing their object identities. The transaction
preserves cell stores, formulas, comments, styles, geometry, unknown fields,
and unrelated package members; it returns an exact-source patch and inverse.
See the focused [`move_table` example](../litchi-numbers/examples/move_table.rs).

Populated sheets can be duplicated adjacent to their source with preserved
sheet settings, drawable order, table names and positions, local formula
dependency graphs, and independently writable table and text-box storage.
Unsupported drawable graphs and cross-table dependency edges fail before the
editor is modified. See `duplicate_numbers_sheet`.

Populated tables can be duplicated with independent storage and CalculationEngine
owner families. Formula hosts, table UUID references, dependency tiles, package
UUID registries, and cross-component references are remapped transactionally;
unsupported advanced dependency state is rejected without modifying the editor.
See `duplicate_numbers_table`.

Pages body tables and Keynote slide tables can use the same native graph clone
without first saving or opening a template document. Pages inserts the copied
inline attachment at an explicit UTF-16 body position; Keynote appends the
copied drawable to the slide and offsets it for direct selection. Both retain
independent cell storage and formula-owner state. See `duplicate_pages_table`
and `duplicate_keynote_table`.

Ordinary sheet-owned text boxes expose UTF-16 text replacement, geometry,
hyperlink, lock, aspect-ratio, and accessibility-property updates. Their
four-object private graph—shape, title and caption stand-ins, and writable
storage—can be duplicated with fresh document-component UUID mappings and
deleted with inbound-reference checks. Clones use Numbers' native 10-point
offset and independent storage. Text, geometry, property, and duplicate/delete
restoration cycles are byte-exact on decompressed native IWA members. Whole-text
replacement also retains Numbers' explicit index-zero drop-cap sentinel. See
`inspect_numbers_text_boxes`, `edit_numbers_text_box`,
`edit_numbers_text_box_geometry`, `edit_numbers_text_box_properties`,
`duplicate_numbers_text_box`, and `remove_numbers_text_box`.

String, formula, rich-text, comment, error, and other `TableDataList` refcount
updates retain raw root or segmented entries and segment references. Tile rows
and row-header buckets likewise patch only the affected buffers, counts, and
ranges. Unknown fields remain in place, malformed duplicate identifiers fail
transactionally, and a real text-cell create/clear cycle restores every
decompressed Numbers member exactly.

The compatibility `NumbersEditor` comment surface decodes root or segmented
`COMMENT_STORAGE` lists and exposes their text, creation date, author, replies,
and storage UUID. Selector-first `litchi_numbers::Package` projections cover
ID-free root-comment and admitted direct-reply reads; direct-reply creation,
replacement, removal, and graph ownership remain compatibility-host scope.
The compatibility editor's root and direct-reply create/update/delete
operations preserve the cell value and style, retain app metadata on in-place
edits, use copy-on-write for shared threads, maintain list refcounts and BNC
flags, and reclaim unreferenced roots, replies, authors, and segment objects.
Comment identifier `1` remains reserved,
matching Numbers' native first key of `2`. Numbers silently rejects fabricated
cell-comment authors, so creation reuses the first registered native annotation
author and fails before mutation when a real package's author storage is still
empty; creating and saving one comment in Numbers primes that identity.
Comment-only empty cells and tables that did not yet have a comment list are
created transactionally. Adding that list patches only the nested
`DataStore.commentStorageTable` reference instead of re-encoding the table
model, so unknown table-model and data-store extensions remain intact. The
`edit_numbers_comment_reply` example exercises the remaining
compatibility-host direct-reply thread layer.

Direct drawable comments use the shared `IWorkDrawableCommentEditor` across
Pages, Numbers, and Keynote protobufs. It resolves every supported nesting of
`TSD.DrawableArchive.comment`, preserves creation date, author, replies, and
message metadata during updates, isolates shared comments with copy-on-write,
and removes orphaned reply graphs on delete. Direct replies support ordered
read, create, update, and delete; native-style root and reply copy-on-write
keeps shared threads isolated, preserves stable storage UUIDs for logical
updates, and reclaims obsolete roots, replies, and generated authors.
`PagesEditor` restricts these
operations to drawables reachable through the document, floating-drawable,
z-order, template, and metadata graphs; `NumbersEditor` restricts them to a
sheet's unique `drawable_infos` ownership; `KeynoteEditor` restricts them to a
slide's `owned_drawables` list. The `inspect_drawable_comments`,
`edit_drawable_comment`, `edit_drawable_comment_reply`, and
`edit_numbers_drawable_comment` examples expose the application-independent
and sheet-scoped APIs. The root editor and raw-ID host-specific direct-drawable
comment methods (including raw-ID reply methods) are explicitly deprecated
compatibility surfaces. Their signatures and behavior are unchanged, and
migration-host examples retain `#![allow(deprecated)]` so existing workflows
continue to compile. Use focused format-semantic comment owners where they
exist; direct-drawable comment ownership has no drop-in focused replacement
yet. Focused text comment/reply APIs and generated/raw modules are not part of
this deprecation slice.
Nested comment references and comment-storage text/UUID fields are patched at
the protobuf wire level, retaining unknown Apple fields byte-for-byte; the
the `litchi-iwa-archive` `compare_iwa_packages` example compares decompressed
object streams independently of Snappy block choices.

Pages sections can be appended by cloning a reachable section's layout and
template references at the current UTF-16 body end, then removed without
deleting body text. Both operations patch only the repeated section-boundary
record and retain unknown protobuf fields. Body insertion keeps the mandatory
initial section boundary at index zero. Selector-first section-scoped text
read, UTF-16 span replacement, whole-value update, and clear now live in
`litchi-pages::Package`; see `litchi-pages/examples/edit_section_text.rs`.
Use the focused `litchi-pages/examples/edit_section_text.rs` example and its
selector-first `litchi_pages::Package` transaction; there is no umbrella
`edit_pages_section_text` example in this migration host.
For a rooted exact source with one unambiguous native body storage, the changed
transaction excludes native U+0004 separators and dependent footnote or
inline-object anchors, preserves unrelated raw records, and publishes only
after a retained-limit reopen and semantic readback. Global whole-body editing
is a single-section convenience so it cannot silently orphan section graphs.
Private Buffa lazy views validate known body-graph fields while raw records
remain the unknown-content preservation authority.
The legacy `PagesEditor` raw-ID section-text methods have been retired; use the
focused package transaction above. Changed nested-`Index.zip` packages still
require the migration host because their physical section-text ownership has
not yet crossed this boundary.
Changed no-root/fallback bodies are likewise unsupported until their physical
ownership has an explicit preservation-safe mutation boundary.

`litchi-pages::document_settings` owns document body/header/footer visibility,
facing-page layout, automatic hyphenation, ligatures, and footnote formatter
settings. Its composite `Package` transaction preserves optional native
presence, validates known formatter values, requires an exact source for a
changed edit, invalidates dependent layout caches and stale previews, and
supports exact inverse patches. See
`litchi-pages/examples/edit_document_settings.rs`.
Page dimensions, margins, scale, orientation, and vertical-layout flags belong
to the selector-free, document-wide `litchi-pages::Package` transaction shown
above and in `litchi-pages/examples/edit_page_layout.rs`. That transaction
retains unknown source bytes, rejects malformed selected fields rather than
normalizing them, requires an exact source for changes, invalidates dependent
layout caches, removes stale root previews, and supports exact inverse patches.
Settings stored directly on a section belong to the selector-first aggregate
`litchi_pages::section::settings` transaction shown above and in
`litchi-pages/examples/edit_section_settings.rs`. It preserves the native
presence of the name, four Boolean flags, and three pagination fields, proves
template dependencies, rewrites only the selected Section component, preserves
ViewState, the derived layout cache, and root previews exactly, and supports
exact inverse patches. The focused section-name and pagination APIs remain
ergonomic facades;
future pagination values remain typed `Unknown` variants and starting page
numbers use a validated non-zero type. The migration host no longer exposes
aggregate raw-ID section-settings reads or writes.
Reachable `TP.PlaceholderArchive` and `TSWP.ShapeInfoArchive` drawables expose
their owned text storages in stable object order. Text-box content supports
UTF-16 range replacement, whole-value update, and clear operations; detached
drawables and shared storage ownership are rejected before mutation. The
`edit_pages_text_box` example exercises the same guarded API on native files.
Body-anchored ordinary text boxes can also be duplicated with independent
shape, storage, attachment, title, and caption objects, then deleted with
orphan checks. Duplication inserts both the body U+FFFC attachment anchor and
the document z-order reference, keeps the attachment table in UTF-16 index
order, advances the package object-identifier high-water mark, and allocates
document-component UUID mappings. The clone is offset by 12 points in each
axis so it remains independently selectable in Pages. Deletion reverses those
registrations and safely releases a contiguous identifier suffix. Clone/delete
cycles restore every decompressed IWA member exactly, including unknown fields,
package metadata, and reference metadata. See `duplicate_pages_text_box` and
`remove_pages_text_box`.
Reachable ordinary text boxes also expose typed position, size, geometry flags,
and rotation in degrees. Optional zero-valued fields retain their raw presence semantics;
updates preserve unknown fields nested inside the geometry, point, and size
messages. See `edit_pages_text_box_geometry`.
Hyperlink URL, lock state, aspect-ratio constraint, and accessibility
description are likewise readable and writable without normalizing unrelated
drawable fields. Pages exposes aspect-ratio constraints for anchored text
boxes but can disable its Lock control for that placement mode. See
`edit_pages_text_box_properties`.

Keynote show dimensions and playback flags are owned by
`litchi_keynote::show::{Settings, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` and the focused `litchi_keynote::Package::{show_settings,
edit_show_settings, apply_show_settings}` transaction. It uses semantic
settings only, preserves unknown content in an exact source, and produces an
immutable committed package plus an exact-source inverse patch. Read-only
legacy nested-`Index.zip` input remains preservable and can be streamed back
unchanged with `Package::write_to`; a changed settings commit deliberately
returns `litchi_keynote::show::Error::UnsupportedSource` rather than normalize
or rebuild that legacy layout. See
`litchi-keynote/examples/edit_show_settings.rs` for safe distinct-output,
sibling-temporary, no-clobber publication.

Slide skip state and navigator name remain bounded compatibility-editor
operations. The focused transition API exposes typed None, Dissolve, Magic
Move, and future effects plus twist, mosaic, bounce, Magic Move fading, timing
curves, text delivery, motion blur, travel distance, animation color, seeds,
detail, curve theme names, and right-to-left writing direction. Validated raw
records preserve unknown nested extensions at the slide, transition,
attributes, and animation-attributes levels.
Soundtrack playback settings are owned by the focused
`litchi_keynote::soundtrack` package transaction, which edits only optional
mode and finite `0.0..=1.0` volume values and preserves future mode values.
It does not expose, rebuild, or reorder soundtrack media. The IWA migration
host deliberately retains media-item and asset CRUD, including their native
references and ordering; that scope is separate from playback settings.
The host can still enumerate slide-owned `TSWP.ShapeInfoArchive` storages for
legacy compatibility, including title, body, and ordinary text-box
classification. Its generic storage operations are raw-ID migration surfaces,
not the supported title/body/notes API; use `litchi-keynote::Package` for those
ordinary semantic edits. Duplicate drawable or cross-slide storage ownership
is rejected before host mutation.
Ordinary text boxes can also be duplicated with independent shape, title and
caption stand-ins, and storage objects, then deleted with inbound-reference
checks. Both slide ownership and z-order lists are patched, the clone is offset
by 10 points in each axis, and the slide-component UUID map and package object
high-water mark advance together. A duplicate/delete cycle restores every
decompressed IWA member byte-for-byte. See `duplicate_keynote_text_box` and
`remove_keynote_text_box`; `inspect_keynote_text_boxes` prints the ordinary
text-box indexes accepted by both examples.
Position, size, geometry flags, and rotation in degrees are readable and writable through
the same ownership guard. The bounded patcher preserves unknown fields at every
nested geometry level and rejects duplicate or malformed scalar encodings. See
`edit_keynote_text_box_geometry`.
Shared drawable properties use the same ownership guard and wire-preserving
mutation path. Keynote exposes Lock for ordinary text boxes but can disable its
aspect-ratio control when the text box height is auto-sized. See
`edit_keynote_text_box_properties`.

Text storages shared by all three applications expose explicit language runs
through `TextLanguage`, `TextLanguageTag`, and `TextPosition`. Setters accept
only scalar UTF-16 boundaries, preserve unknown table and entry fields, and
coalesce redundant adjacent runs. Scratch-created text boxes inherit the
document language, matching real Pages, Numbers, and Keynote output. Individual
nonzero boundaries or the complete language table can be deleted without
changing text or sibling formatting. See `edit_iwork_text_language` and
`inspect_iwork_text_styles`.
The same storages expose native hyperlinks as `TextHyperlink` values with
strict `TextRange`, `TextHyperlinkTarget`, and `TextHyperlinkId` types. Create
rejects overlaps with hyperlinks or other smart fields, update retains the
smart-field identity, and delete reclaims the owned object and package
identifier suffix. Unknown fields in the storage table, individual boundaries,
and hyperlink payload survive edits. Web URLs, `mailto:` links, and Keynote
targets such as `?slide=next` are represented losslessly. The Pages, Numbers,
and Keynote text-box editors provide ownership-checked wrappers; see
`edit_iwork_text_hyperlink` and `inspect_iwork_text_styles`.
Native Date & Time fields use `TextDateTimeField` and lossless typed formatter
settings. Existing text can be attached to a field, or `insert_*_date_time_field`
can atomically insert caller-supplied localized display text and its smart-field
object. The crate deliberately does not emulate Apple's locale formatter: the
display text is explicit while the native date/time pattern string, locale,
date/time styles, refresh plan, update flag, and Apple-reference-date instant
remain structured. The pattern string is bounded at the focused owner boundary;
its grammar is not validated.
Deletion retains visible text and reclaims the owned field object; ordinary text
replacement also reclaims orphaned fields. Pages body and text-box, Numbers
sheet text-box, and Keynote slide text-box wrappers enforce document ownership.
See `create_iwork_date_time_field` and `inspect_iwork_date_time_fields`.
Native textual number attachments are represented by `TextNumberAttachment`.
Insertion atomically adds the U+FFFC placeholder, point-table entry, object, and
indexed reference; deletion reverses that graph and ordinary text replacement
reclaims orphaned objects. Typed settings distinguish page number, page count,
footnote mark, and unknown future kinds while preserving optional string and
number-format metadata. Pages exposes body, reachable header/footer, and text-box
wrappers. The shared low-level reader also decodes native Keynote slide-number
storage without claiming that ordinary Numbers or Keynote text boxes evaluate
page numbers. See `create_pages_number_attachments`,
`create_pages_text_box_number_attachment`, and
`inspect_iwork_number_attachments`.
Pages body text additionally exposes native ranged bookmarks as `TextBookmark`
values. `TextBookmarkSettings` carries an optional validated name and a lossless
visible/hidden value; create and update reject empty, overlapping, out-of-bounds,
or surrogate-splitting ranges. Deletion and text replacement reclaim the owned
`TSWP.BookmarkFieldArchive`, while unknown bookmark and boundary fields survive
updates. See `create_pages_bookmark` and `inspect_pages_bookmarks`.
Native plain highlights use `TextHighlight`, `TextHighlightId`, and the same
strict `TextRange` boundaries. Creation builds the complete native annotation
graph with author, timestamp, and UUID metadata; range updates retain object
identity and unknown wire fields; deletion reclaims the owned empty comment
storage and generated author when unused. Plain highlights and ranged comments
are classified independently while sharing one strictly validated native range
table. The ownership-checked Pages, Numbers, and Keynote wrappers are
demonstrated by `edit_iwork_text_highlight`; `inspect_iwork_text_styles` reports
both highlight IDs and ranges.
Ranged comments use `TextComment`, `TextCommentId`, and the nonempty
`TextCommentBody` newtype. Create emits the complete native type-2013/type-3056
graph used by all three applications. Update can move the range and replace the
body without changing the annotation ID, timestamp, author, storage UUID, or
reply thread. Ordered direct replies use `TextCommentReply`,
`TextCommentReplyId`, and `TextCommentReplyBody`; create, read, update, and
delete retain reply identity, timestamp, author, storage UUID, ordering, and
unknown wire fields. Deletion validates exclusive ownership and reclaims the
target storage and generated author when unused. Root deletion reclaims the
entire thread. Unknown fields at the table, boundary, annotation, root, and
reply-storage levels survive updates. See `edit_iwork_text_comment` and the
three scratch text-box creation examples.
Slide move, duplicate, and delete operations likewise rewrite only the nested
slide-tree ownership list. Existing raw `TSP.Reference` payloads are reused,
so extensions inside the show, slide tree, and individual references survive;
duplicate/delete and move/restore cycles return the original decompressed IWA
members byte-for-byte. Slide duplication remaps references in slide,
placeholder, note, text-storage, shape, and slide-node payloads through
schema-bounded wire paths, preserving unknown fields at every ancestor and
inside the remapped references themselves.

### Keynote slide-table persisted lock state uses the focused package owner

Persisted slide-table lock reads and exact lock-state transactions are owned by
`litchi-keynote::Package::{slide_table_lock_state,
edit_slide_table_lock_state, apply_slide_table_lock_state}`. The public
`State` and transaction values are archive-free; native IDs, Archive/ZIP/member
values, raw wire views, and generated/Prost/Buffa types stay private. Strict
`table_info_codec` admission and prepared lock rewriting preserve unrelated
fields, admit exact lock and unlock transitions, and fail closed for malformed,
ambiguous, or unsupported sources.

Raw Keynote title, persisted-sort, persisted-lock, and physical `Sort Now`
configuration methods and calls are retired from the legacy editor host. The
focused `litchi-keynote::Package` owns
`execute_slide_table_sort_order(SlideSelector, TableSelector)` and
`execute_slide_table_sort_order_to_rows(SlideSelector, TableSelector,
RowRange)`, while the old raw-ID `apply_*` methods remain deprecated
declarations for source compatibility and the boundary checker rejects
production calls to them. There is no lock or physical-sort fallback for an
exact package: focused errors propagate. A guarded source-built-only writer is
retained for `KeynoteDocumentBuilder` graphs carrying the explicit
`Application/Litchi/Blank/Wide` marker; unmarked exact packages cannot fall
back after focused-owner refusal. Physical row/value reading, formulas, rich
cells, and broader table compatibility remain outside the focused owner.

Physical row planning is archive-free and shared through the hidden common
`RowPermutation` primitive. The focused owner admits only the canonical
type-6001 table-model route and explicitly proven tile, data-list, header, UID,
and empty pre-BNC sentinel shapes. It validates rule arity and scalar domains,
uses source offsets as the final key for deterministic stable duplicate
ordering, and builds its inverse with bounded staging. The private adapter
borrows BNC cell views while planning and rejects hostile row, column, and
key-product dimensions before allocation. Formula/error cells, rich text,
comments, merges, filters, groups, categories, pivots, spills, conditional
styles, hidden/non-positional state, imported/provenance data, non-empty
stroke, cross-tile/cross-bucket movement, unknown mutable fields, and other
unproven row-affine state fail closed atomically.

The current strict owner rejects the app-authored Keynote 14.4 probe because
model field 39 identifies an unowned conditional-style CalculationEngine
dependency graph. A pre-hardening Computer Use run is
recorded in ADR 0008 as external exploratory evidence only; it does not
certify current-owner native E3/E4 acceptance. The checked-in native fixture
has no table and the checked-in evidence test records hashes without launching
Keynote. Native acceptance remains pending.

The scoped gates are 17/17 for the lock codec, 9/9 for focused Keynote lock
integration, 30/30 for the IWA slide-table suite, four migrated Keynote example
checks, and 625/625 boundary tests with passing `py_compile`. Strict
Keynote/protos checks and Clippy passed, as did both isolated fuzz checks with
15 codec seeds and eight lifecycle command seeds. Scoped title/sort/lock audits
are empty; the live full checker still reports only three unrelated untracked
Pages table-lock violations. No full-workspace-green claim follows.

One actual temporary-driver attempt used
`/private/tmp/wave100-native-table-headers-source.key` (500,128 bytes,
SHA-256
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`, 86
members). Strict Package lock read at slide 0/table 0 returned `InvalidSource`
before mutation; source bytes/hash stayed exact, no candidate or inverse was
emitted, and no UI run occurred. Native lock mutation/save/reopen acceptance is
withheld.

The operation ledger is conservative logical accounting, not Package or
SourceCatalog cache/decompressed-Archive, allocator/RSS, or codec-internal
allocator telemetry; no zero-copy or RSS claim follows. Wave105 closes no
crate, manifest dependency, dependency edge, or migration debt. The
`litchi-iwa -> litchi-keynote` edge, all 13 ordered debts (including 014, 015,
016, and 017), generated/Prost/Buffa owners, and the IWA monolith deletion gate
remain unchanged.

### Numbers FormulaArchive extraction uses a bounded neutral projection

The legacy `litchi-iwa::numbers::table_extractor::TableDataExtractor` now
retains strict bounded owned FormulaArchive wire bytes and renders through
the neutral `numbers_formula_codec` scalar and compatibility visitors. The
sidecar owns its preservation bytes; it is not a zero-copy view. Production
generated `tsce::FormulaArchive` decoding and generated AST construction are
removed from this extractor path, while the former generated renderer remains
available only as a `cfg(test)` differential oracle.

Canonical but incomplete postfix programs intentionally retain legacy output:
operand-less negation renders `=FORMULA()` and surplus expressions select the
final expression. Malformed wire, duplicated known fields, and noncanonical
known values remain rejected.

`FormulaOwnerDependencies`, category maps, and name maps remain permissive
generated best-effort compatibility products. Formula authoring, formula-cache
refresh, physical formula mutation, cloning, merge handling, and dependency
shifting remain generated compatibility responsibilities. This section does
not claim a full Numbers formula migration or generated-free physical table
editing.

One `ProjectionBudget` spans reference-map census, sidecar admission,
repeated renders, and table extraction; formula decode-report fields, work,
and text usage are merged into that budget. The ledger is conservative logical
accounting, not Package/cache or decompressed-Archive, allocator, RSS, or
codec-internal telemetry. No zero-copy, allocator/RSS, latency, or
package-wide performance claim follows.

Wave106 scoped gates are 38/38 focused extractor tests; green `litchi-iwa`
library check, no-run, and strict Clippy; a passing formula fuzz binary check
and strict Clippy with a 14-seed smoke run completed for 100 runs; a dedicated
FormulaArchive extractor boundary audit with no findings; and 643/643 full
boundary unit tests with passing `py_compile`.

Read-only native evidence uses the Numbers-created
`/private/tmp/wave106-numbers-formula-source.numbers` (136,591 bytes,
SHA-256
`81fe99b6647e370d1b1663c703500f1fac239111ae7b4bea862f74803c94208c`, 43
members). It contains one 22-by-7 `Table 1`; the migrated extractor read six
materialized cells and reported `=(B3+C3)` at zero-based `(1,3)` and
`=SUM(B4:C4)` at `(2,3)`. No mutation, candidate, inverse, save/reopen, or UI
acceptance claim follows.

The legacy host still owns formula authoring/cache/mutation and related
physical table, tile, dependency, clone, and merge paths. Wave106 closes no
crate, manifest dependency, edge, or migration debt; all 13 ordered debts,
the generated/Prost/Buffa owners, and the IWA monolith deletion gate remain
unchanged.

### Pages body-table names use the focused package owner

Pages body-table name reads and rename transactions are owned by the
selector-first `litchi_pages::Package` APIs
`body_table_name`, `edit_body_table_name`, and `apply_body_table_name`. The
public `table::name::Name` and transaction values are archive-free. The strict
`table_model_discovery_codec` field-8 projection/prepared rewrite preserves
unknown canonical fields and groups while rejecting malformed, duplicate,
noncanonical, and wrong-wire input.

Raw `PagesEditor` rename mutation is retired and the example uses the focused
package owner. `PagesEditor::tables()` retains a legacy generated read-only
name-listing compatibility path outside focused owner admission; it has no
mutation fallback. Physical Pages table/content, storage, formula, and other
compatibility duties remain in `litchi-iwa`.

Wave107 gates are 22/22 for focused Pages body-table-name integration, 38/38
for host Pages table tests, 9/9 for the focused codec, and 565/565 for the
full `litchi-iwa-protos` library suite. Strict Pages owner/test Clippy passed;
the `edit_pages_table` example check passed; lifecycle fuzz check and strict
target Clippy passed with 10 command seeds. Boundary unit tests passed
648/648 and live name HOST/FACADE audits are empty. The full checker remains
blocked only by an unrelated pre-existing untracked Pages table-lock file.
These are scoped gates, not a full-workspace result.

The name owner proves the selected model's unique current component and
locator and rejects unknown, external, data, ambiguous, and root-map metadata
routes. Unrelated UUID bits and component assignments remain
opaque-preserved, not independently verifiable here.

No positive native Pages name/rename acceptance is claimed. The only available
body-table source,
`/private/tmp/wave104-pages-native-source.pages` (108,776 bytes, SHA-256
`997509fda639f5dcdabd4546c392b9ebdc3d8c7e9c1d967f35b9f2a0aca38359`, 43
members), was rejected by strict selector read with
`InvalidSource { path: Table { table: 0 } }`. Its source bytes remained exact;
no candidate or inverse was produced and no UI run occurred.

The operation ledger is conservative logical accounting. It does not measure
Package/SourceCatalog caches, decompressed Archives, process allocator/RSS,
or codec-internal allocator telemetry; no zero-copy, allocator/RSS, or
package-wide performance claim follows. The `litchi-iwa -> litchi-pages` edge,
debt 017 and all 13 ordered debts, generated/Prost/Buffa ownership, and the
IWA monolith deletion gate remain unchanged.

### Keynote slide-table dimensions use the focused package owner

Selector-first Keynote slide-table dimension reads and transactions are owned
by `litchi_keynote::Package::{slide_table_dimension_size,
edit_slide_table_dimension_size, apply_slide_table_dimension_size}`. The
public `slide::table::dimension::{Dimension, Points, Size}` values and
transaction values are archive-free. Native identifiers, Archive/ZIP/member
data, raw wire views, and generated/Prost/Buffa values remain private.

The owner atomically rewrites the selected `HeaderStorageBucket` size and the
matching `TableInfo` drawable geometry. Strict role, ArchiveInfo/FieldInfo,
metadata, global-inbound, and persisted-lock checks run before publication.
One conservative logical budget spans source admission, candidate reopen and
semantic reread, locality, exact inverse, and stale/conflicting patch
rejection. `litchi-iwa` retains physical resize/geometry, rows/cells,
storage, formulas, tiles, and broader compatibility helpers.

The global physical census enforces UUID-pair uniqueness and rejects all-zero
identifiers before publication; unrelated metadata assignments remain outside
this focused route.

Raw persisted-dimension Keynote editor methods and wrappers are retired, as is
the obsolete Keynote appearance bridge. Two remaining legacy geometry/remove
and archive-name lookups use `KeynoteObjectCatalog`; physical helpers and
legacy graph compatibility remain in the migration host.

Wave108 scoped gates are a passing Keynote owner library check and strict
library Clippy; 12/12 focused dimension tests with strict test Clippy; 30/30
IWA slide-table host tests; passing `create_keynote_table` and
`list_keynote_tables` example checks; and a passing lifecycle fuzz-target
check plus strict target Clippy with 10 command-only seeds. Focused dimension
boundary checks passed 5/5, the full boundary unit suite passed 653/653 with
passing `py_compile`, and live facade/host audits were empty. The top-level
checker still reports only three unrelated pre-existing untracked Pages
table-lock findings; these are scoped results and do not indicate full
workspace health.

No positive native Keynote dimension acceptance is claimed. The one authorized
read used `/private/tmp/wave100-native-table-headers-source.key` (500,128
bytes, SHA-256
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`, 86
members). Strict Row selection returned `UnsupportedDependency`; the source
remained byte-exact, no candidate or inverse was emitted, no UI run occurred,
and the repository was not changed. Native dimension mutation, save, reopen,
and UI acceptance evidence is withheld.

The operation budget is a conservative logical envelope, not telemetry for
`Package`/`SourceCatalog` caches, decompressed Archives, the process
allocator, RSS, or codec-internal allocation behavior. No zero-copy,
allocator/RSS, or package-wide performance claim follows. The
`litchi-iwa -> litchi-keynote` edge, all 13 ordered debts (including 017),
generated/Prost/Buffa ownership, migration-host role, and the IWA monolith
deletion gate remain unchanged. Wave108 closes no crate, manifest dependency,
edge, debt, host owner, generated owner, or monolith gate.

## Embedded media

Single-file packages and in-memory bytes expose their `Data/*` members. The
metadata-backed editor retains stable data identifiers and reference counts,
and patches only the digest and materialized length fields at the protobuf wire
level so unknown Apple extensions remain byte-exact.

`MediaAsset::media_type` uses the compact, archive-free
`litchi_iwa_common::media::Type` value. It is a one-byte copyable classification
for image, video, audio, PDF, and unknown assets; package discovery, metadata,
limits, and replacement validation remain owned by this crate.

```rust
use litchi_iwa::IWorkMediaEditor;

let mut media = IWorkMediaEditor::open("input.key")?;
let asset = media
    .assets()
    .iter()
    .find(|asset| asset.is_materialized())
    .expect("materialized asset")
    .clone();
let replacement = std::fs::read("replacement.jpg")?;
media.replace(asset.data_identifier, &replacement)?;
media.save("updated.key")?;
# Ok::<(), Box<dyn std::error::Error>>(())
```

`PagesEditor::section_media_assets`, `NumbersEditor::sheet_media_assets`, and
`KeynoteEditor::slide_media_assets` scope discovery through the authoritative
object-reference graph. Their `replace_media` methods reject identifiers that
are not reachable from the application document root. `remove_unreferenced`
removes only records absent from component records, message data references,
and `DataMetadataMap`; referenced deletion is rejected transactionally.

Pre-iWork '13 single-file documents that wrap a directory-style bundle are
normalized by this legacy host on import. Their IWA components, operation log,
media, previews, and metadata remain available to its compatibility APIs. This
does not make normalized legacy sources generally editable through
the concrete package crates: a changed selector-first transaction may return
`UnsupportedSource` until it has an explicit preservation-safe owner.

## Build Requirements

The companion `litchi-iwa-protos` crate compiles the raw protobuf definitions
via `prost-build`; the `protoc` compiler must be available on `PATH`:

- Debian / Ubuntu: `apt install protobuf-compiler`
- macOS (Homebrew): `brew install protobuf`

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.

## Wave109 amendment: Keynote slide-table names remain a focused package owner

Selector-first litchi_keynote::Package owns archive-free slide-table name
reads and exact name transactions through
Package::{slide_table_name, edit_slide_table_name, apply_slide_table_name}.
The public value is slide::table::name::Name. The strict
table_model_discovery_codec field-8 prepared rewrite preserves unknown
canonical fields/groups and rejects malformed, duplicate-known, noncanonical,
or wrong-wire input before publication; native identifiers, Archive/ZIP/member
data, wire views, and generated/Prost/Buffa values remain private.

The admitted route is a rooted, unique slide -> table-info -> model ->
storage/name route. Strict storage-route, role, ArchiveInfo,
metadata/current-component, UUID, and global-inbound authority checks precede
staging, with fail-closed behavior for unsupported, ambiguous, aliased,
locked, malformed, or otherwise unproven routes. Focused transaction tests
verify source and unrelated-byte preservation, semantic readback/locality, and
exact inverse restoration including canonical preview state. The raw Keynote
rename mutation path is retired; no fallback remains. IWA continues to own
physical rows/cells, storage, formulas, tiles, and broader table compatibility.

Wave109 verification: 18/18 focused slide-table-name tests; strict
litchi-keynote library and test Clippy PASS; 30/30 IWA host slide-table tests
and litchi-iwa library check PASS; slide-table-name fuzz-target check and
strict target Clippy PASS with 10 command-only hex seeds; 662/662 boundary
unit tests; py_compile PASS; full checker PASS; live
FACADE=[]; RESOURCE=[]; HOST=[].

No positive native Keynote name/rename acceptance is claimed. The strict read
of /private/tmp/wave100-native-table-headers-source.key (500,128 bytes,
SHA-256 47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b,
86 members) returned InvalidSource. The source remained exact; no candidate
or inverse was produced and no UI run occurred. Native mutation, save,
reopen, and UI acceptance evidence is withheld.

One conservative logical operation ledger is recorded; Package cache,
decompressed-Archive, OS/process allocator, and RSS telemetry are
unobservable. No zero-copy, allocator/RSS, or package-wide performance claim
is made. This amendment changes no crate, manifest, dependency edge, debt,
host owner, generated owner, or monolith gate. The topology remains 64
packages, 239 internal dependency declarations, one migration host, and 13
ordered debts; litchi-iwa -> litchi-keynote/debt 014, all other edges/debts,
generated/Prost/Buffa ownership, and the IWA monolith deletion gate remain
unchanged.

## Wave110 amendment: Numbers dimension bucket wire is generated-free

The private Numbers table-dimension storage adapter now reads and rewrites
type-6006 header buckets through the neutral `table_dimension_codec`, not the
generated `TST.HeaderStorageBucket` or nested header types. Reads validate the
whole bucket under finite limits; writes use a prepared plan, inspect its
requirements, execute once, preserve untouched raw fields/order, and publish
atomically. Duplicate/out-of-range records, invalid sizes, mismatched row
hashes, duplicate/zero row-bucket references, and known table/storage role
aliases fail closed. External storage references, incorrect bucket counts,
row/column storage aliasing, and row records outside their assigned 65,536-row
slot are rejected as well.

This is a private wire migration only. The surrounding route still uses a
generated `TableModelArchive` for legacy selection; it is not a new public
package owner, a package-wide shared-bucket authority proof, or a raw-ID API
retirement. No native/UI acceptance is claimed.

Wave110 verification passed 58/58 focused codec tests, protos check/strict
Clippy, five IWA regressions, IWA library check/strict scoped Clippy, 667/667
boundary tests, `py_compile`, and the live storage-codec audit. The full
checker's only failures were three unrelated findings from the pre-existing
untracked Pages table-lock file. Limits are conservative logical envelopes,
not package-cache, decompressed-Archive, ZIP/Snappy, allocator, RSS, or
zero-copy telemetry. The 64-package/239-internal-dependency/13-debt topology,
all dependency edges, and the IWA monolith deletion gate remain unchanged.

## Wave111 Keynote movie-transform ownership

The selector-first `litchi_keynote::Package` movie-geometry transaction now
keeps `MovieGeometry` as the archive-free position/display-size value and adds
archive-free `MovieTransform` (finite angle plus supported reflection) and
`MovieFlipAxis`. `Package::{slide_movie_geometry, edit_slide_movie_geometry,
apply_slide_movie_geometry}` stages geometry and transform atomically, with
exact-source fingerprints, no-op/conflict and inverse behavior, strict
candidate reread, preview invalidation, and locality checks. Native IDs,
archives, media assets, and unrelated flag bits remain private.

The strict `keynote_movie_geometry_codec` handles optional native flags and
angle fields through finite prepared limits and one execution, preserving
unknown fields/groups and unrelated flag bits while rejecting duplicate,
wrong-wire, malformed, and non-finite input. The focused package and codec
suites are 27/27 and 10/10. The four raw Keynote host geometry/restore/flip
entry points and their fallback are retired; typed host bridges delegate to
the focused package and propagate its errors. Private physical helpers remain
only for unrelated media lifecycle, offset, property, and graph work.

Both the low-level `keynote_movie_geometry_codec` and package-level
`keynote_slide_movie_geometry` fuzz targets passed isolated cargo checks and
strict target Clippy, and the bounded 32-run smoke passed. The package-level
corpus contains nine hand-authored command seeds, with eight new command seeds
across the package and low-level targets; no native package artifact is
implied. The final completion ratchet passed its focused 8/8 and full 670/670
boundary checks; `py_compile` passed and live audits were empty. The full
checker reported only the three unrelated findings from the pre-existing
untracked Pages table-lock file. Both fuzz targets passed isolated checks,
strict target Clippy, and the bounded 32-run smoke; the package-level corpus
contains nine hand-authored command seeds, with eight new seeds across both
targets and no native package artifact. A broader `litchi-keynote` library
sweep was 152/153, with the sole `soundtrack_order` failure unrelated to
Wave111. No Wave111 native movie source exists:
tracked `basic.key` contains no `.mov`, and historical Wave87 artifacts are
absent. No Wave111 candidate/inverse/UI/save/reopen acceptance is claimed;
the historical Wave87 geometry record and its strict normalized reread
`InvalidSource` remain unchanged.

This is a bounded ownership extension, not a monolith-exit gate. The
64-package/239-internal-dependency/13-debt topology, all dependency edges,
`litchi-iwa -> litchi-keynote`, generated/Prost/Buffa ownership, and the IWA
monolith gate remain unchanged. Resource limits are conservative logical
envelopes; package caches, decompressed Archives, ZIP/Snappy buffers,
codec-internal allocation, the process allocator, RSS, and zero-copy behavior
are not directly measured.

## Wave117 Keynote chart-legend visibility ownership

Selector-first `litchi_keynote::Package` now owns the effective visibility of
an existing Keynote slide-chart legend through
`Package::{slide_chart_legend_visible, edit_slide_chart_legend,
apply_slide_chart_legend}`. Select slides and charts by semantic position or
exact visible name. The archive-free transaction values are
`ChartLegendVisibility{Edit,Patch,Commit,Diagnostics,Error,LimitKind}`;
native object IDs, component names, archives, generated messages, and Buffa
views remain private. An absent native field reads as effective visibility
`false` while its presence is preserved in the source-backed implementation.

The focused owner performs exact-source no-op/change, inverse, conflict,
candidate-readback, root-preview invalidation, and package-wide locality
checks. Its private
`litchi-iwa-protos::keynote_chart_legend_codec` uses strict preflight followed
by a lazy Buffa projection and prepared source-preserving rewrite of the chart
non-style field 20. Unknown fields, groups, and ordering remain untouched.

The raw-ID `KeynoteEditor::{slide_chart_legend_visible,
set_slide_chart_legend_visible}` methods are retired with no fallback. The
legacy host still owns Keynote legend fill/frame/font/shadow/stroke, chart
creation and duplication/removal, and broader graph work; Pages and Numbers
continue to use their host legend operations.

Wave117 verification records 13/13 codec tests, 10/10 focused package tests,
the allocation-mapping unit, strict library/test Clippy, all-target checks,
696/696 boundary-policy tests, and a 1,000-run nightly fuzz smoke over 12
checked-in corpus seeds. Computer Use opened the 46,022-byte focused output
without repair, showed the selected chart's Legend checkbox off, saved and
reopened a 154,754-byte native copy with Legend still off, and the focused API
reverse-read and reproduced the native copy byte-for-byte. The focused and
native SHA-256 values are
`2ecf1327971c206e4017168cfd5fac86b0954ae32a0f2f0ba894f3b6b84c5acd`
and `b82af597b4469b056b559cde678db508fca13d7ec725633bdd82749fab560adf`.
These are disposable acceptance artifacts, not checked-in fixtures or a
complete monolith-exit claim.

## Wave119 debt-005 archive/core boundary

The neutral archive primitives used by the migration host are now reached
through the archive owner's doc-hidden `litchi_iwa_archive::iwa` route. The
inspection examples and the low-level compatibility example use that route
rather than naming `litchi_iwa_core` directly; native archive identities and
preservation behavior are unchanged. The `litchi-iwa` manifest no longer
carries a direct dependency on `litchi-iwa-core`; the archive owner remains
responsible for the physical archive/core boundary.

This retires ordered migration debt 005 only. The workspace topology is 64
packages, 237 internal dependency declarations, 226 canonical edges, and one
migration host, with the remaining ordered debts `[1, 2, 4, 8, 10, 12, 13,
14, 15, 16, 17]`. The IWA monolith deletion gate, generated/Prost/Buffa
ownership, and the other migration-host edges remain open. This amendment
makes no semantic-mutation or performance claim.

Computer Use opened the canonical Pages, Numbers, and Keynote fixtures,
created native copies, closed them, and reopened the copies without repair or
recovery UI. Pages retained its three text/date markers, Numbers retained the
fixture marker and value `42`, and Keynote retained its title/body/date
markers. The format-owned `Package::save` APIs then republished those
native-normalized copies byte-for-byte: Pages was 96,407 bytes with SHA-256
`d321bde90824664eb6122690eacd30441aa5d4d329c655b808e1b420a78e6bb5`,
Numbers was 135,985 bytes with SHA-256
`1e23a5b36e3f11bc0de2b11de37488c4c981f13f7335ef415e63eccb2adedc18`,
and Keynote was 499,981 bytes with SHA-256
`a720f3a1dbe32070a1c72bc710621747b9c305261b86c4ade1879b6a3eadaf02`.
These were disposable preservation artifacts, not checked-in fixtures.

## Wave120 bounded TableDataList text projection

The private generic text registry no longer eagerly constructs generated Prost
`TST.TableDataList` or `TST.TableDataListSegment` values for message types
6005, 6201, and 6011. It now routes those payloads through the existing strict
`numbers_table_cell_storage_codec`, whose handwritten validation and private
Buffa lazy-view parity keep the caller-owned archive bytes authoritative.

Only validated, non-empty cell strings are copied into a fallibly allocated
staging vector, and that vector is published only after the complete payload
passes its byte, field, work, nesting, reference, and text limits. This is a
private extraction-path migration: editor mutation paths and other generated
decoders remain open, and it changes no public API, dependency edge, ordered
migration debt, or monolith-deletion gate.

## 2026-08-31 Keynote slide-table persisted sort transaction hardening

The existing selector-first Keynote persisted field-44 sort transaction now
reuses the shared `slide_table_core` authority for canonical admission, its
bounded operation-local budget, and its locality checks. Ambiguity in archive
roles, references, types, or wire framing fails closed before publication.
Exact aggregate-only producer metadata and current, unversioned in-package
cross-component edges remain admissible; partial route metadata, dangling or
duplicate edges, and foreign inbound ownership remain rejected. The
source-preserving rewrite continues to preserve previews, field 45,
unknown fields, and untouched archive entries, while retaining exact no-op,
inverse, and conflict behavior.

This is persisted-configuration hardening only. At the 2026-08-31 checkpoint,
physical `Sort Now` and `RowRange`-based row execution were still host-owned;
the later focused `litchi-keynote` owner now owns those physical-sort
transactions. The legacy host retains only its explicitly source-built
compatibility route and broader cells/table-storage work; exact focused-owner
refusals are terminal and do not fall back to a host physical executor. The
amendment makes no native semantic acceptance, performance, fuzz-exhaustiveness,
broader Keynote authoring, or monolith-deletion claim. The authoritative
topology remains 64 workspace packages, 237 internal dependency declarations,
226 canonical edges, and 11 ordered migration debts
`[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, with one migration host; no
migration debt is retired and no host exits.

## 2026-09-01 Numbers table relocation focused owner

Physical relocation of an existing Numbers table now belongs to
`litchi_numbers::table::relocation::transaction::{Commit, Diagnostics, Edit,
Error, LimitKind, Patch, Path}` through
`litchi_numbers::Package::{edit_table_relocation, move_table,
apply_table_relocation}` (`apply_table_move` remains a compatibility spelling).
The contract is
selector-first (`SheetSelector` source and destination plus a source-sheet
`TableSelector`) and exposes no native IDs or raw archive values. Same-sheet
moves are exact no-ops; changed moves provide exact-source, conflict-checked
patches and inverses with locality checks while preserving table content.

The public legacy `NumbersEditor::move_table` entry point and its host example
are retired. Exact snapshots use the focused semantic `litchi-numbers`
transaction, and any focused refusal is terminal. A crate-private physical
bridge remains only for populated-sheet duplication of historical source-built
packages whose cell storage is outside the current semantic projection; that
bridge uses the doc-hidden, selector-first admission seam in `litchi-numbers`,
the same rewrite and verification engine, and legacy candidate readback before
publication. Unsupported graphs remain fail-closed. This note supersedes the
older migration prose that described physical table moves as generally
host-owned. It is a bounded ownership transfer only: no migration debt,
dependency edge, host, or monolith-deletion claim changes. The current topology
remains 64 packages, 238 internal declarations, 227 canonical edges, 11
development-only edges, 11 ordered debts, and one migration host.
