# litchi-iwa

Apple iWork archive reader and writer for `.pages`, `.numbers`, and `.key` files.

## Overview

`litchi-iwa` reads Apple iWork bundles using their IWA (iWork Archive) layout: a ZIP container holding Snappy-compressed, protobuf-encoded object streams along with media assets and metadata. It exposes a unified `Document` API that handles all three iWork applications, plus lower-level access to archives, the object reference graph, and structured content (tables, slides, sections).

## Usage

```toml
[dependencies]
litchi-iwa = "0.0.1"
```

```rust
use litchi_iwa::Document;

let doc = Document::open("document.pages")?;
let text = doc.text()?;
let stats = doc.stats();
println!("objects: {}", stats.total_objects);

let structured = doc.extract_structured_data()?;
println!("{}", structured.summary());
# Ok::<(), litchi_iwa::Error>(())
```

## Features

- Parse Pages, Numbers, and Keynote bundles from a path or in-memory bytes
- Build independent Pages, Numbers, and Keynote packages from typed IWA objects
  with no bundled template
- Snappy decompression and protobuf decoding of `.iwa` streams
- Text extraction across all iWork applications
- Structured-data extraction: tables (with CSV export), slides, sections
- Metadata-backed media discovery, extraction, replacement, and guarded cleanup
- Typed cross-suite image, movie, and audio-property read/write for hyperlinks,
  locking, aspect-ratio locking, and accessibility descriptions with lossless
  unknown-field preservation
- Metadata-preserving IWA object/message create, read, update, and delete operations
- Snappy IWA serialization and deterministic package rewriting
- Transactional package-entry and IWA-component updates with atomic saves
- Legacy nested `Index.zip` bundle import with byte-preserved assets and
  explicit password-protected document rejection
- Semantic editors for Numbers sheets/tables/cells/formulas, Pages
  body/header/footer/text-box text, and Keynote slides/placeholders/text boxes/speaker notes
- Native-style ordinary shape duplication across Pages, Numbers, and Keynote
  with independent rich-text storage, fresh UUID mappings, preserved opaque
  fields, and app-specific selection offsets
- Native Pages, Numbers, and Keynote chart CRUD, including source-built inline
  data charts and duplicate operations with fresh private graphs, preserved
  editable data, theme-preset registration, UUIDs, and native placement offsets
- Typed copy-on-write text-box paragraph alignment, native line-spacing modes,
  atomic before/after spacing, first-line/left/right indentation, and ordered
  left/center/right/decimal tab stops with leaders across Pages, Numbers, and Keynote
- Typed BCP 47 text-language run CRUD at validated UTF-16 scalar boundaries,
  including automatic-language sentinels and lossless boundary deletion
- Native cross-suite text-hyperlink CRUD with typed nonempty UTF-16 ranges,
  lossless web, mail, and Keynote navigation targets, and owned-object cleanup
- Native cross-suite Date & Time smart-field CRUD with typed ICU formats,
  locale identifiers, formatter styles, refresh plans, and Apple-reference instants
- Native Pages page-number/page-count attachment CRUD with typed kinds, exact
  U+FFFC placement, lossless number metadata, and header/footer support
- Native Pages body-bookmark CRUD with typed names, lossless visibility values,
  strict UTF-16 ranges, and owned bookmark-field cleanup
- Native cross-suite plain-text highlight CRUD with typed nonempty UTF-16 ranges,
  lossless table mutation and owned annotation cleanup
- Native cross-suite ranged text-comment and ordered direct-reply CRUD with
  nonempty typed bodies, stable IDs and metadata, and scratch-package creation
- Lossless Pages document body/header/footer visibility, facing-page layout,
  automatic hyphenation, and ligature options
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
- Native Keynote build-in/build-out object CRUD with typed On Click / After Transition /
  With Previous / After Previous timing, typed Rotate / Scale / Opacity / Move actions,
  editable Bézier motion paths and custom timing curves, typed Blink / Bounce / Flip / Jiggle / Pop / Pulse
  emphasis actions, typed Keyboard / Shimmer / Skid / Swoosh / Trace build-in/build-out
  effects with native direction models, wire-preserving
  move/reorder operations, validated raw CRUD for unmapped native build parameters,
  typed transition effects, acceleration and text delivery, semantic transition
  deletion through Keynote's native `none` state, lossless effect-specific
  parameters, component UUIDs,
  and slide-node cache maintenance
- Typed direct-drawable comment CRUD with Pages document-reachability,
  Numbers sheet-ownership, and Keynote slide-ownership guards
- Native Numbers cell-comment and direct-reply CRUD with table-list refcounts,
  copy-on-write threads, annotation authors, dates, and UUIDs

## Semantic editing

Native charts can be created directly from typed `ChartData` with
`add_body_chart`, `add_sheet_chart`, or `add_slide_chart`; no source table or
template package is required. Their corresponding `duplicate_body_chart`,
`duplicate_sheet_chart`, and `duplicate_slide_chart` methods retain the source
data and opaque chart fields while giving the new chart independent inline data,
private styles, preset registration, UUIDs, ownership, and native placement.
See `create_*_chart` and `duplicate_*_chart` in `examples/` for runnable
file-to-file workflows.

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

All source-built and existing direct drawables also expose native Arrange
stacking through `body_drawable_order`, `sheet_drawable_order`, and
`slide_drawable_order`. Each list runs back-to-front; its setter requires an
exact permutation, while `move_*_drawable` accepts typed
`DrawableLayerMove::{ToBack, Backward, Forward, ToFront}` commands. See the
`create_*_stacked_shapes` examples for complete scratch-file workflows.

Ordinary source-built and existing shapes expose the native Flip buttons via
`DrawableFlipAxis::{Horizontal, Vertical}` and `flip_body_shape`,
`flip_sheet_shape`, or `flip_slide_shape`. Each operation updates native
geometry while retaining unrelated fields; see `create_*_flipped_shape` for
scratch-file examples.

File-backed images use the same typed command through `flip_body_image`,
`flip_sheet_image`, or `flip_slide_image`, preserving their embedded asset,
adjustments, and metadata; see `create_*_flipped_image` for scratch-file
examples.

### Create Pages documents from scratch

```rust
use litchi_iwa::pages::PagesEditor;

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

let mut pages = PagesEditor::create_with_text("Quarterly revenue\n")?;
let first_anchor = pages.body_text()?.encode_utf16().count();
let table = pages.add_table(first_anchor, "Revenue", 4, 3)?;
pages.set_table_cell(
    table.model_object_id,
    0,
    0,
    CellValue::Text("Quarter".to_owned()),
)?;
pages.rename_table(table.model_object_id, "Revenue by Quarter")?;
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
# Ok::<(), litchi_iwa::Error>(())
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
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapePreset};

let body = "Quarterly report";
let mut pages = PagesEditor::create_with_text(body)?;
let shape = pages.add_body_shape(
    body.encode_utf16().count(),
    "A fully editable shape",
    DrawablePoint { x: 180.0, y: 240.0 },
    DrawableSize { width: 300.0, height: 150.0 },
    ShapePreset::RightArrow,
)?;
pages.set_body_shape_text(shape.drawable_object_id, "Updated")?;
pages.set_body_shape_preset(shape.drawable_object_id, ShapePreset::DoubleArrow)?;
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
use litchi_iwa::shapes::{DrawablePoint, LineEndpoint, LineEndpoints};

let mut pages = PagesEditor::create_with_text("Built without a template")?;
let line = pages.add_body_line_with_endpoints(
    pages.body_text()?.encode_utf16().count(),
    DrawablePoint { x: 180.0, y: 240.0 },
    DrawablePoint { x: 480.0, y: 390.0 },
    LineEndpoints::new(LineEndpoint::OpenCircle, LineEndpoint::FilledArrow),
)?;
pages.set_body_line_segment(
    line.drawable_object_id,
    DrawablePoint { x: 96.0, y: 180.0 },
    DrawablePoint { x: 456.0, y: 180.0 },
)?;
assert_eq!(
    pages.body_line_endpoints(line.drawable_object_id)?.end,
    LineEndpoint::FilledArrow,
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
    DrawablePoint { x: 96.0, y: 144.0 },
    DrawableSize { width: 300.0, height: 225.0 },
)?;
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
use litchi_iwa::pages::{PagesEditor, PagesMovieOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

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
        DrawablePoint { x: 96.0, y: 144.0 },
        DrawableSize { width: 320.0, height: 180.0 },
        Duration::from_secs(8),
    ),
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
use litchi_iwa::pages::{PagesAudioOptions, PagesEditor};
use litchi_iwa::shapes::DrawablePoint;

let body = "Interview notes";
let audio = fs::read("interview.aiff")?;
let mut pages = PagesEditor::create_with_text(body)?;
let source = pages.add_body_audio(
    body.encode_utf16().count(),
    "interview.aiff",
    &audio,
    PagesAudioOptions::new(
        DrawablePoint { x: 180.0, y: 240.0 },
        Duration::from_secs(30),
    ),
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
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

let mut numbers = NumbersDocumentBuilder::new()
    .sheet_name("Forecast")
    .table_name("Revenue")
    .build()?;
let sheet_id = numbers.sheets()?[0].object_id;
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
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapePreset};

let mut numbers = NumbersDocumentBuilder::new().build()?;
let sheet_id = numbers.sheets()?[0].object_id;
let shape = numbers.add_sheet_shape(
    sheet_id,
    "A fully editable shape",
    DrawablePoint { x: 420.0, y: 300.0 },
    DrawableSize { width: 300.0, height: 150.0 },
    ShapePreset::RightArrow,
)?;
numbers.set_sheet_shape_text(sheet_id, shape.drawable_object_id, "Updated")?;
numbers.set_sheet_shape_preset(sheet_id, shape.drawable_object_id, ShapePreset::DoubleArrow)?;
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
use litchi_iwa::shapes::{DrawablePoint, LineEndpoint, LineEndpoints};

let mut numbers = NumbersDocumentBuilder::new().build()?;
let sheet_id = numbers.sheets()?[0].object_id;
let line = numbers.add_sheet_line_with_endpoints(
    sheet_id,
    DrawablePoint { x: 420.0, y: 300.0 },
    DrawablePoint { x: 720.0, y: 450.0 },
    LineEndpoints::new(LineEndpoint::FilledCircle, LineEndpoint::SimpleArrow),
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
    LineEndpoint::FilledCircle,
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
let sheet_id = numbers.sheets()?[0].object_id;
let source = numbers.add_sheet_image(
    sheet_id,
    "chart.png",
    &image,
    DrawablePoint { x: 420.0, y: 180.0 },
    DrawableSize { width: 320.0, height: 240.0 },
)?;
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
let sheet_id = numbers.sheets()?[0].object_id;
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
let sheet_id = numbers.sheets()?[0].object_id;
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

### Create Keynote presentations from scratch

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
keynote.set_slide_number_visible(1, true)?;
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
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapePreset};

let mut keynote = KeynoteDocumentBuilder::new().build()?;
let shape = keynote.add_slide_shape(
    0,
    "A fully editable shape",
    DrawablePoint { x: 720.0, y: 660.0 },
    DrawableSize { width: 480.0, height: 240.0 },
    ShapePreset::RightArrow,
)?;
keynote.set_slide_shape_text(0, shape.drawable_object_id, "Updated")?;
keynote.set_slide_shape_preset(0, shape.drawable_object_id, ShapePreset::DoubleArrow)?;
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
use litchi_iwa::shapes::{DrawablePoint, LineEndpoint, LineEndpoints};

let mut keynote = KeynoteDocumentBuilder::new().build()?;
let line = keynote.add_slide_line_with_endpoints(
    0,
    DrawablePoint { x: 720.0, y: 660.0 },
    DrawablePoint { x: 1_200.0, y: 900.0 },
    LineEndpoints::new(LineEndpoint::OpenSquare, LineEndpoint::FilledDiamond),
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
    LineEndpoint::FilledDiamond,
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
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

let image = fs::read("chart.png")?;
let mut keynote = KeynoteDocumentBuilder::new().build()?;
let source = keynote.add_slide_image(
    0,
    "chart.png",
    &image,
    DrawablePoint { x: 704.0, y: 284.0 },
    DrawableSize { width: 512.0, height: 512.0 },
)?;
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
use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteSlideMovieOptions};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

let movie = fs::read("demo.mov")?;
let poster = fs::read("demo-poster.png")?;
let mut keynote = KeynoteDocumentBuilder::new().build()?;
let source = keynote.add_slide_movie(
    0,
    "demo.mov",
    &movie,
    "demo-poster.png",
    &poster,
    KeynoteSlideMovieOptions::new(
        DrawablePoint { x: 640.0, y: 360.0 },
        DrawableSize { width: 640.0, height: 360.0 },
        Duration::from_secs(8),
    ),
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
use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteSlideAudioOptions};
use litchi_iwa::shapes::DrawablePoint;

let audio = fs::read("narration.aiff")?;
let mut keynote = KeynoteDocumentBuilder::new().build()?;
let source = keynote.add_slide_audio(
    0,
    "narration.aiff",
    &audio,
    KeynoteSlideAudioOptions::new(
        DrawablePoint { x: 960.0, y: 540.0 },
        Duration::from_secs(12),
    ),
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

The same media APIs expose `MediaPlaybackSettings` for typed trim boundaries,
poster position, repeat mode, and `MediaVolume`. Update the returned settings
and write them through the matching `*_playback_settings` method (for example,
`set_body_movie_playback_settings` or
`set_slide_audio_playback_settings`). The update preserves unrelated and
unknown movie-archive fields; `MediaVolume::new` rejects invalid levels, and
`MediaLoopMode::Unknown` allows a newer native repeat value to round-trip.

### Edit existing documents

```rust
use litchi_iwa::numbers::{
    CellValue, FormulaAxisReference, FormulaCellReference, FormulaExpression, NumbersEditor,
    NumbersTableHeaderCount,
    NumbersTableTitleSettings,
};
use litchi_iwa::pages::{PagesDocumentOptions, PagesEditor};
use litchi_iwa::keynote::{
    KeynoteBuildSettings, KeynoteBuildStart, KeynoteEditor, KeynoteFlipDirection,
    KeynoteHorizontalBuildDirection, KeynoteKeyboardDirection, KeynoteRotationDirection,
    KeynoteSlideTextRole, KeynoteSwooshDirection,
};

let mut numbers = NumbersEditor::open("input.numbers")?;
let table = numbers.tables()?.remove(0);
let mut headers = numbers.table_header_settings(table.object_id)?;
headers.header_rows = Some(NumbersTableHeaderCount::TWO);
headers.footer_rows = Some(NumbersTableHeaderCount::ONE);
headers.header_rows_frozen = Some(true);
numbers.set_table_header_settings(table.object_id, headers)?;
numbers.set_table_title_settings(
    table.object_id,
    NumbersTableTitleSettings {
        visible: Some(true),
        outlined: Some(false),
    },
)?;
numbers.set_cell(table.object_id, 1, 2, CellValue::Number(42.0))?;
// Existing rich-text cells use the same call. Their TSWP formatting storage is
// retained, and shared payloads are isolated with copy-on-write.
numbers.set_cell(table.object_id, 1, 3, CellValue::Text("Revised".into()))?;
numbers.set_cell_comment(table.object_id, 1, 3, "Check this value")?;
let _comment = numbers.cell_comment(table.object_id, 1, 3)?;
let reply_id = numbers.add_cell_comment_reply(table.object_id, 1, 3, "Looks good")?;
let reply_id = numbers.set_cell_comment_reply(
    table.object_id,
    1,
    3,
    reply_id,
    "Verified",
)?;
numbers.remove_cell_comment_reply(table.object_id, 1, 3, reply_id)?;
numbers.clear_cell_comment(table.object_id, 1, 3)?;
numbers.set_formula(
    table.object_id,
    2,
    2,
    FormulaExpression::function(
        "SUM",
        [FormulaExpression::Number(1.0), FormulaExpression::Number(2.0)],
    ),
)?;
numbers.set_formula(
    table.object_id,
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
    table.object_id,
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
numbers.resize_table(table.object_id, 30, 10)?;
numbers.rename_table(table.object_id, "Inventory")?;
let original_sheet_id = numbers.sheets()?[0].object_id;
let copied_sheet = numbers.duplicate_sheet(original_sheet_id)?;
numbers.remove_sheet(copied_sheet.object_id)?;
let new_sheet = numbers.add_empty_sheet("Archive")?;
let new_table = numbers.add_empty_table(new_sheet.object_id, "Log", 100, 6)?;
numbers.move_table(table.object_id, new_sheet.object_id)?;
numbers.move_table(table.object_id, original_sheet_id)?;
numbers.move_sheet(new_sheet.index, 0)?;
numbers.remove_table(new_table.object_id)?;
numbers.remove_sheet(new_sheet.object_id)?;
if let Some(sheet) = numbers.sheets()?.first()
    && let Some(text_box) = numbers.sheet_text_boxes(sheet.object_id)?.first()
{
    numbers.set_sheet_text_box_text(
        sheet.object_id,
        text_box.drawable_object_id,
        "Updated text box",
    )?;
    let geometry =
        numbers.sheet_text_box_geometry(sheet.object_id, text_box.drawable_object_id)?;
    numbers.set_sheet_text_box_geometry(
        sheet.object_id,
        text_box.drawable_object_id,
        geometry,
    )?;
    let properties =
        numbers.sheet_text_box_properties(sheet.object_id, text_box.drawable_object_id)?;
    numbers.set_sheet_text_box_properties(
        sheet.object_id,
        text_box.drawable_object_id,
        properties,
    )?;
    numbers.set_sheet_drawable_comment(
        sheet.object_id,
        text_box.drawable_object_id,
        "Review this text box",
    )?;
    let _comment =
        numbers.sheet_drawable_comment(sheet.object_id, text_box.drawable_object_id)?;
    numbers.clear_sheet_drawable_comment(sheet.object_id, text_box.drawable_object_id)?;
    let copy = numbers.duplicate_sheet_text_box(
        sheet.object_id,
        text_box.drawable_object_id,
        "Independent copy",
    )?;
    numbers.remove_sheet_text_box(sheet.object_id, copy.drawable_object_id)?;
}
numbers.save("updated.numbers")?;

let mut pages = PagesEditor::open("input.pages")?;
let mut document_options = pages.document_options()?;
document_options.facing_pages = Some(true);
document_options.automatic_hyphenation = Some(true);
document_options.ligatures_enabled = Some(false);
pages.set_document_options(document_options)?;
let section_id = pages.sections()[0].object_id;
pages.set_section_text(section_id, "Updated body")?;
let first_header = pages
    .header_footers()?
    .into_iter()
    .find(|region| matches!(region.kind, litchi_iwa::pages::PagesHeaderFooterKind::Header))
    .expect("document header");
pages.set_header_footer_text(first_header.storage.object_id, "Quarterly report")?;
pages.set_section_name(section_id, Some("Executive summary"))?;
let mut section_settings = pages.section_settings(section_id)?;
section_settings.inherit_previous_header_footer = Some(false);
section_settings.first_page_hides_header_footer = Some(true);
section_settings.start = Some(litchi_iwa::pages::PagesSectionStart::NextPage);
section_settings.page_numbering = Some(
    litchi_iwa::pages::PagesSectionPageNumbering::Restart,
);
section_settings.starting_page_number = Some(
    litchi_iwa::pages::PagesPageNumber::new(3)?,
);
pages.set_section_settings(section_id, section_settings)?;
pages.set_section_background(
    section_id,
    litchi_iwa::pages::PagesSectionBackground::Solid(litchi_iwa::pages::PagesRgbaColor {
        red: 1.0,
        green: 0.59,
        blue: 0.55,
        alpha: 1.0,
        color_space: litchi_iwa::pages::PagesRgbColorSpace::Srgb,
    }),
)?;
let inserted = pages.insert_section(section_id, 8, "Methods")?;
pages.remove_section(inserted.object_id)?;
let appended = pages.append_section(section_id, "Appendix")?;
pages.remove_section(appended.object_id)?;
let mut layout = pages.page_layout()?;
layout.top_margin = Some(54.0);
layout.orientation = Some(litchi_iwa::pages::PagesPageOrientation::Portrait);
pages.set_page_layout(layout)?;
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
keynote.set_slide_title(0, "Updated title")?;
keynote.set_slide_body(0, "Updated body")?;
keynote.set_slide_notes(0, "Presenter cue")?;
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
keynote.set_slide_name(0, Some("Opening"))?;
keynote.set_slide_skipped(0, false)?;
keynote.set_slide_number_visible(0, true)?;
let layout = keynote.default_slide_layout()?;
let fresh = keynote.add_slide(layout)?;
keynote.set_slide_title(fresh.index, "New from theme")?;
let mut transition = keynote.slides()?[0].transition.clone().expect("transition");
transition.effect = Some(litchi_iwa::keynote::KeynoteTransitionEffect::Dissolve);
transition.duration = Some(1.5);
transition.custom_parameters.acceleration =
    Some(litchi_iwa::keynote::KeynoteTransitionAcceleration::EaseInOut);
transition.custom_parameters.text_delivery =
    Some(litchi_iwa::keynote::KeynoteTransitionTextDelivery::ByWord);
keynote.set_slide_transition(0, transition)?;
keynote.clear_slide_transition(0)?;
let mut show = keynote.show_settings()?;
show.loop_presentation = Some(true);
show.mode = Some(litchi_iwa::keynote::KeynoteShowMode::SelfPlaying);
keynote.set_show_settings(show)?;
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
    keynote.move_slide_build(0, build.object_id, 0)?;
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
let copy = keynote.duplicate_slide(0)?;
keynote.move_slide(copy.index, 0)?;
keynote.remove_slide(0)?;
keynote.save("updated.key")?;
# Ok::<(), litchi_iwa::Error>(())
```

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

Numbers sheet names (including nested form-based sheets), table names, and
required table-model dimensions are patched at the protobuf wire level.
Unrecognized Apple fields retain their bytes and position, while duplicate
singular fields fail transactionally. Header rows, header columns, footer rows,
freeze toggles, and print-time repeating-header toggles expose their native
optional presence through typed settings; the validated count type matches
Numbers' native 1–5 choices, with `None` representing zero. See
`edit_numbers_table_headers`. Full-table sort rules have typed read, set, and
clear APIs that preserve native rule extensions and reference-tracker metadata.
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
`NumbersEditor::add_empty_table` can also recreate the first native table after
the workbook's last table was removed. It derives the style graph from the
workbook theme, builds independent storage and row/column identities, and
registers a fresh formula owner without relying on a hidden template table.
Populated tables can also move between sheets without changing their object
identity, cell stores, formulas, comments, styles, or geometry. The operation
transfers the original raw drawable reference, rewrites the optional table
parent, and updates both sheets' IWA reference metadata atomically. See
`move_numbers_table`.

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

Numbers cell comments are decoded from root or segmented `COMMENT_STORAGE`
lists and expose their text, creation date, author, replies, and storage UUID.
Root and direct-reply create/update/delete operations preserve the cell value
and style, retain app metadata on in-place edits, use copy-on-write for shared
threads, maintain list refcounts and BNC flags, and reclaim unreferenced roots,
replies, authors, and segment objects. Comment identifier `1` remains reserved,
matching Numbers' native first key of `2`. Numbers silently rejects fabricated
cell-comment authors, so creation reuses the first registered native annotation
author and fails before mutation when a real package's author storage is still
empty; creating and saving one comment in Numbers primes that identity.
Comment-only empty cells and tables that did not yet have a comment list are
created transactionally. Adding that list patches only the nested
`DataStore.commentStorageTable` reference instead of re-encoding the table
model, so unknown table-model and data-store extensions remain intact. The
`edit_numbers_comment` and `edit_numbers_comment_reply` examples exercise both
thread layers.

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
and sheet-scoped APIs.
Nested comment references and comment-storage text/UUID fields are patched at
the protobuf wire level, retaining unknown Apple fields byte-for-byte; the
`compare_iwa_packages` example compares decompressed object streams independently
of Snappy block choices.

Pages sections can be appended by cloning a reachable section's layout and
template references at the current UTF-16 body end, then removed without
deleting body text. Both operations patch only the repeated section-boundary
record and retain unknown protobuf fields. Body insertion keeps the mandatory
initial section boundary at index zero. Section-scoped text supports read,
UTF-16 range replacement, whole-value update, and clear operations without
exposing the native U+0004 separators. Boundary positions and header/footer
locations are refreshed after every edit. Global whole-body replacement is
restricted to single-section documents so it cannot silently orphan section
graphs; use `edit_pages_section_text` for native multi-section files.

Pages page dimensions, margins, scale, orientation, and vertical-layout flags
are also patched directly in the protobuf wire stream. Unknown Apple fields
retain their original bytes and positions; duplicate singular fields, wrong
wire types, and truncated payloads fail transactionally instead of being
normalized by a decode/re-encode cycle. The Document formatter's body, header,
footer, facing-page, automatic-hyphenation, and ligature toggles are likewise
writable through lossless optional settings; absent values retain Pages'
effective defaults. See `edit_pages_document_options`. Page orientation,
facing-page section
starts, and continue/restart numbering behavior use lossless enums; future
native values remain available as typed `Unknown` variants. Starting page
numbers use a validated non-zero type. See `edit_pages_layout` and
`edit_pages_section_pagination`.
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
package metadata, and reference metadata. See `duplicate_pages_text_box`,
`remove_pages_text_box`, and `inspect_pages_text_boxes`.
Reachable ordinary text boxes also expose typed position, size, geometry flags,
and rotation in degrees. Optional zero-valued fields retain their raw presence semantics;
updates preserve unknown fields nested inside the geometry, point, and size
messages. See `edit_pages_text_box_geometry`.
Hyperlink URL, lock state, aspect-ratio constraint, and accessibility
description are likewise readable and writable without normalizing unrelated
drawable fields. Pages exposes aspect-ratio constraints for anchored text
boxes but can disable its Lock control for that placement mode. See
`edit_pages_text_box_properties`.

Keynote show dimensions and playback flags, slide skip state and navigator
name, and modern transition identifiers/timing fields use the same bounded wire
mutations. Transition effects are typed, including native None, Dissolve, and
Magic Move variants plus lossless future identifiers. Clearing an effect emits
Keynote's real `none` representation while retaining start timing and the native
random seed. Native twist, mosaic, bounce, Magic Move fading, timing curve, text
delivery, motion blur, and travel-distance transition parameters are available
through a lossless CRUD model, with typed acceleration and text delivery.
Modern animation color, arbitrary native
timing-curve payloads, random seeds, effect detail, curve theme names, and
right-to-left writing direction are writable as well. Unknown nested transition
extensions remain byte-exact at the slide, transition, transition-attributes,
and animation-attributes levels.
Slide-owned `TSWP.ShapeInfoArchive` storages are enumerated in drawable order
and classified as title, body, or ordinary text boxes. Their content supports
UTF-16 range replacement, whole-value update, and clear operations while
preserving unaffected style and annotation records. Duplicate drawable or
cross-slide storage ownership is rejected before mutation.
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
display text is explicit while the ICU pattern, locale, date/time styles,
refresh plan, update flag, and Apple-reference-date instant remain structured.
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

## Embedded media

Single-file packages and in-memory bytes expose their `Data/*` members. The
metadata-backed editor retains stable data identifiers and reference counts,
and patches only the digest and materialized length fields at the protobuf wire
level so unknown Apple extensions remain byte-exact.

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

## Low-level CRUD

```rust
use litchi_iwa::{IWorkPackage, archive::RawMessage};

let mut package = IWorkPackage::open("document.pages")?;
package.update_archive("Index/Document.iwa", |archive| {
    let object = archive.object_mut(1).expect("document root");
    object.replace_message(0, RawMessage {
        type_: object.messages[0].type_,
        data: object.messages[0].data.clone(),
    })?;
    Ok(())
})?;
package.save("updated.pages")?;
# Ok::<(), litchi_iwa::Error>(())
```

Pre-iWork '13 single-file documents that wrap a directory-style bundle are
expanded to the modern flat package layout on import. Their IWA components,
operation log, media, previews, and metadata remain available under normalized
paths, so semantic and low-level edits use the same APIs as current documents.

## Build Requirements

This crate compiles protobuf definitions via `prost-build`. The `protoc` compiler must be available on `PATH`:

- Debian / Ubuntu: `apt install protobuf-compiler`
- macOS (Homebrew): `brew install protobuf`

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
