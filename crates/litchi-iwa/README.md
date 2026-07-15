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
- Snappy decompression and protobuf decoding of `.iwa` streams
- Text extraction across all iWork applications
- Structured-data extraction: tables (with CSV export), slides, sections
- Metadata-backed media discovery, extraction, replacement, and guarded cleanup
- Metadata-preserving IWA object/message create, read, update, and delete operations
- Snappy IWA serialization and deterministic package rewriting
- Transactional package-entry and IWA-component updates with atomic saves
- Legacy nested `Index.zip` bundle import with byte-preserved assets and
  explicit password-protected document rejection
- Semantic editors for Numbers sheets/tables/cells/formulas, Pages
  body/header/footer/text-box text, and Keynote slides/placeholders/text boxes/speaker notes
- Lossless Pages document body/header/footer visibility, facing-page layout,
  automatic hyphenation, and ligature options
- Typed Numbers table header/footer counts, freeze state, and repeating-header
  settings with lossless optional-field presence
- Typed Keynote theme-layout discovery and fresh empty-slide creation with native
  component registration, speaker notes, slide numbers, and transactional insertion
- Wire-preserving Keynote per-slide number visibility with native placeholder
  ownership and z-order invariants
- Native Keynote build-in/build-out object CRUD with typed On Click / After Transition /
  With Previous / After Previous timing, typed Rotate / Scale / Opacity / Move actions,
  editable Bézier motion paths, typed Blink / Bounce / Flip / Jiggle / Pop / Pulse
  emphasis actions, typed Keyboard / Shimmer / Skid / Swoosh / Trace build-in/build-out
  effects with native direction models, wire-preserving
  move/reorder operations, validated raw CRUD for unmapped native build parameters,
  typed transition acceleration and text delivery, lossless effect-specific
  transition parameters, component UUIDs,
  and slide-node cache maintenance
- Typed direct-drawable comment CRUD with Pages document-reachability,
  Numbers sheet-ownership, and Keynote slide-ownership guards
- Native Numbers cell-comment and direct-reply CRUD with table-list refcounts,
  copy-on-write threads, annotation authors, dates, and UUIDs

## Semantic editing

```rust
use litchi_iwa::numbers::{
    CellValue, FormulaAxisReference, FormulaCellReference, FormulaExpression, NumbersEditor,
    NumbersTableHeaderCount,
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
transition.duration = Some(1.5);
transition.custom_parameters.acceleration =
    Some(litchi_iwa::keynote::KeynoteTransitionAcceleration::EaseInOut);
transition.custom_parameters.text_delivery =
    Some(litchi_iwa::keynote::KeynoteTransitionTextDelivery::ByWord);
keynote.set_slide_transition(0, transition)?;
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
    let mut build_settings = build.settings.clone();
    build_settings.effect = "apple:dissolve character".to_owned();
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
`edit_numbers_table_headers`. Table resizing still updates tiles,
header buckets, stable UID maps, and stroke sidecars as one checked operation;
each existing object is now mutated through bounded wire paths, including the
unpacked UID index arrays and nested UUID records. Grow/shrink restoration is
byte-exact while unknown fields remain attached to retained rows, headers,
UUIDs, and stroke-layer references.
Workbook sheet ordering and standard or form-sheet table ownership lists reuse
the original raw `TSP.Reference` payloads, preserving extensions inside each
reference; newly appended references are removed byte-exactly on rollback or
create/delete cycles.
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
name, and modern transition strings/timing fields use the same bounded wire
mutations. Native twist, mosaic, bounce, Magic Move fading, timing curve, text
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
