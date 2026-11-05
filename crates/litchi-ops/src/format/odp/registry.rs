//! The ODP write registry: one table per element kind, layered on the read
//! tables of [`super::spec`] (record 0278).

use super::spec::{
    CELL_PROPS, CHARACTER_KEYS, FRAME_PROPS, GROUP_PROPS, NOTES_PROPS, PARAGRAPH_PROPS, RUN_PROPS,
    SHAPE_PROPS, SLIDE_PROPS, TABLE_PROPS,
};
use crate::Code;
use crate::node::spec::{GapSpec, PropSpec, Subject};
use crate::value::ValueType;

/// The setters of the ODP write registry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Setter {
    /// A story's, paragraph's, run's, cell's or notes' text.
    Text,
    /// `draw:name` of an object or page.
    Name,
    /// `svg:x`.
    X,
    /// `svg:y`.
    Y,
    /// `svg:width`.
    Width,
    /// `svg:height`.
    Height,
    /// Bold, through a derived automatic style (record 0308).
    Bold,
    /// Italic.
    Italic,
    /// Single underline.
    Underline,
    /// Font size.
    Size,
    /// Font colour.
    Color,
    /// Font family.
    Font,
}

const fn same_name(left: &str, right: &str) -> bool {
    let left = left.as_bytes();
    let right = right.as_bytes();
    if left.len() != right.len() {
        return false;
    }
    let mut index = 0;
    while index < left.len() {
        if !left[index].eq_ignore_ascii_case(&right[index]) {
            return false;
        }
        index += 1;
    }
    true
}

/// The registry entries of an element: its readable properties, the
/// writable ones marked with their setter (`names` selects which), and a
/// write-only `text` entry appended when `text` is set.
const fn write_entries<const N: usize>(
    readable: &[PropSpec],
    names: &[&str],
) -> [crate::registry::Entry<Setter>; N] {
    use crate::registry::{Access as WriteAccess, Constraint, Entry};

    let mut text_spec = PropSpec::new("text", ValueType::Text);
    text_spec.access.get = false;
    let text_entry = Entry {
        spec: text_spec,
        access: WriteAccess::Set(Setter::Text),
        constraint: Constraint::None,
        clear: true,
    };
    let mut entries = [text_entry; N];
    let mut index = 0;
    while index < readable.len() {
        entries[index] = Entry::read(readable[index]);
        let name = readable[index].name;
        let mut writable = false;
        let mut candidate = 0;
        while candidate < names.len() {
            if same_name(names[candidate], name) {
                writable = true;
            }
            candidate += 1;
        }
        if writable {
            let setter = if same_name(name, "text") {
                Setter::Text
            } else if same_name(name, "name") {
                Setter::Name
            } else if same_name(name, "x") {
                Setter::X
            } else if same_name(name, "y") {
                Setter::Y
            } else if same_name(name, "width") {
                Setter::Width
            } else if same_name(name, "height") {
                Setter::Height
            } else if same_name(name, "bold") {
                Setter::Bold
            } else if same_name(name, "italic") {
                Setter::Italic
            } else if same_name(name, "underline") {
                Setter::Underline
            } else if same_name(name, "size") {
                Setter::Size
            } else if same_name(name, "color") {
                Setter::Color
            } else {
                Setter::Font
            };
            entries[index].access = WriteAccess::Set(setter);
            if same_name(name, "text") {
                entries[index].clear = true;
            }
            if matches!(setter, Setter::Width | Setter::Height) {
                entries[index].constraint = Constraint::Length {
                    min: crate::value::Length::from_emu(0),
                    max: crate::value::Length::from_emu(i64::MAX),
                };
            }
            if matches!(setter, Setter::Name) {
                entries[index].constraint = Constraint::NotEmpty;
            }
        }
        index += 1;
    }
    entries
}

const GEOMETRY: &[&str] = &["name", "x", "y", "width", "height"];

/// A shape's geometry and the character formatting of its whole text
/// (record 0316).
const SHAPE_WRITABLE: &[&str] = &[
    "name",
    "x",
    "y",
    "width",
    "height",
    "bold",
    "italic",
    "underline",
    "size",
    "color",
    "font",
];

/// Notes text and the character formatting of the whole notes story.
const NOTES_WRITABLE: &[&str] = &[
    "text",
    "bold",
    "italic",
    "underline",
    "size",
    "color",
    "font",
];

const SLIDE_WRITE_COUNT: usize = SLIDE_PROPS.len();
const SHAPE_WRITE_COUNT: usize = SHAPE_PROPS.len() + 1;
const FRAME_WRITE_COUNT: usize = FRAME_PROPS.len();
const TABLE_WRITE_COUNT: usize = TABLE_PROPS.len();
const GROUP_WRITE_COUNT: usize = GROUP_PROPS.len();
const PARAGRAPH_WRITE_COUNT: usize = PARAGRAPH_PROPS.len() + 1;
const RUN_WRITE_COUNT: usize = RUN_PROPS.len() + 1;
const CELL_WRITE_COUNT: usize = CELL_PROPS.len() + 1;
const NOTES_WRITE_COUNT: usize = NOTES_PROPS.len();

pub(crate) const SLIDE_WRITE_TABLE: crate::registry::Table<Setter, SLIDE_WRITE_COUNT, 0> =
    crate::registry::Table::new(write_entries(SLIDE_PROPS, &["name"]), []);
pub(crate) const SHAPE_WRITE_TABLE: crate::registry::Table<Setter, SHAPE_WRITE_COUNT, 0> =
    crate::registry::Table::new(write_entries(SHAPE_PROPS, SHAPE_WRITABLE), []);
pub(crate) const FRAME_WRITE_TABLE: crate::registry::Table<Setter, FRAME_WRITE_COUNT, 0> =
    crate::registry::Table::new(write_entries(FRAME_PROPS, GEOMETRY), []);
pub(crate) const TABLE_WRITE_TABLE: crate::registry::Table<Setter, TABLE_WRITE_COUNT, 0> =
    crate::registry::Table::new(write_entries(TABLE_PROPS, GEOMETRY), []);
pub(crate) const GROUP_WRITE_TABLE: crate::registry::Table<Setter, GROUP_WRITE_COUNT, 0> =
    crate::registry::Table::new(write_entries(GROUP_PROPS, &["name"]), []);
pub(crate) const PARAGRAPH_WRITE_TABLE: crate::registry::Table<Setter, PARAGRAPH_WRITE_COUNT, 0> =
    crate::registry::Table::new(write_entries(PARAGRAPH_PROPS, CHARACTER_KEYS), []);
pub(crate) const RUN_WRITE_TABLE: crate::registry::Table<Setter, RUN_WRITE_COUNT, 0> =
    crate::registry::Table::new(write_entries(RUN_PROPS, CHARACTER_KEYS), []);
pub(crate) const CELL_WRITE_TABLE: crate::registry::Table<Setter, CELL_WRITE_COUNT, 0> =
    crate::registry::Table::new(write_entries(CELL_PROPS, CHARACTER_KEYS), []);
pub(crate) const NOTES_WRITE_TABLE: crate::registry::Table<Setter, NOTES_WRITE_COUNT, 0> =
    crate::registry::Table::new(write_entries(NOTES_PROPS, NOTES_WRITABLE), []);

pub(crate) const SLIDE_CAPABILITY_PROPS: &[PropSpec] = &SLIDE_WRITE_TABLE.readable();
pub(crate) const SHAPE_CAPABILITY_PROPS: &[PropSpec] = &SHAPE_WRITE_TABLE.readable();
pub(crate) const FRAME_CAPABILITY_PROPS: &[PropSpec] = &FRAME_WRITE_TABLE.readable();
pub(crate) const TABLE_CAPABILITY_PROPS: &[PropSpec] = &TABLE_WRITE_TABLE.readable();
pub(crate) const GROUP_CAPABILITY_PROPS: &[PropSpec] = &GROUP_WRITE_TABLE.readable();
pub(crate) const PARAGRAPH_CAPABILITY_PROPS: &[PropSpec] = &PARAGRAPH_WRITE_TABLE.readable();
pub(crate) const RUN_CAPABILITY_PROPS: &[PropSpec] = &RUN_WRITE_TABLE.readable();
pub(crate) const CELL_CAPABILITY_PROPS: &[PropSpec] = &CELL_WRITE_TABLE.readable();
pub(crate) const NOTES_CAPABILITY_PROPS: &[PropSpec] = &NOTES_WRITE_TABLE.readable();

/// The capability property table of an element, by canonical name.
pub(crate) fn capability_props(name: &str) -> Option<&'static [PropSpec]> {
    Some(match name {
        "slide" => SLIDE_CAPABILITY_PROPS,
        "shape" | "placeholder" => SHAPE_CAPABILITY_PROPS,
        "picture" | "chart" | "ole" | "media" => FRAME_CAPABILITY_PROPS,
        "table" => TABLE_CAPABILITY_PROPS,
        "group" | "connector" => GROUP_CAPABILITY_PROPS,
        "paragraph" => PARAGRAPH_CAPABILITY_PROPS,
        "run" => RUN_CAPABILITY_PROPS,
        "tc" => CELL_CAPABILITY_PROPS,
        "notes" => NOTES_CAPABILITY_PROPS,
        _ => return None,
    })
}

/// The canonical names of elements with a set handler.
pub(crate) const SETTABLE: &[&str] = &[
    "slide",
    "shape",
    "placeholder",
    "picture",
    "table",
    "chart",
    "ole",
    "media",
    "group",
    "connector",
    "paragraph",
    "run",
    "tc",
    "notes",
    "hyperlink",
];

/// The canonical names of elements `add` creates.
pub(crate) const ADDABLE: &[&str] = &["slide", "shape", "paragraph", "hyperlink", "linebreak"];

/// The canonical names of elements `remove` deletes.
pub(crate) const REMOVABLE: &[&str] = &[
    "slide",
    "shape",
    "placeholder",
    "picture",
    "table",
    "chart",
    "ole",
    "media",
    "group",
    "connector",
    "paragraph",
    "run",
    "hyperlink",
    "linebreak",
    "video",
    "audio",
    "equation",
];

const fn operation_gap(
    verb: &'static str,
    element: &'static str,
    context: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb,
            element,
            context,
        },
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

/// The ODP write contexts the adapter refuses (record 0278 §6).
pub(crate) static WRITE_GAPS: [GapSpec; 10] = [
    operation_gap(
        "set",
        "run",
        "field-run",
        "M6 ODP field editing",
        "page-number, date and other field runs are computed by the producer and refuse text",
    ),
    operation_gap(
        "set",
        "shape",
        "character-format",
        "0316 ODP inherited shape formatting",
        "bold, italic, underline, size, colour and font are set on paragraphs, runs (links and fields included), and on a shape, table cell or notes story as a whole (each paragraph and every run that overrides the value); reads resolve the node's style chain through the common styles of styles.xml but not a shape's graphic or presentation style, and text sets write plain paragraphs in the first paragraph's style",
    ),
    operation_gap(
        "set",
        "shape",
        "non-paragraph-story",
        "M6 ODP rich text boxes",
        "a whole-story text set refuses a text box that holds drawing objects or other non-paragraph children; set paragraphs or runs instead",
    ),
    operation_gap(
        "add",
        "shape",
        "positioned-insert",
        "M6 ODP z-order insertion",
        "new text boxes are appended at the end of the slide's objects; move them afterwards",
    ),
    operation_gap(
        "add",
        "shape",
        "picture-table-chart",
        "M6 ODP media and table creation",
        "only text boxes are added; pictures, tables, charts and media need package resources the scene splice does not author",
    ),
    operation_gap(
        "move",
        "paragraph",
        "story-reorder",
        "M6 ODP paragraph ordering",
        "paragraphs, runs, rows and cells are not moved or swapped; slides and top-level slide objects are",
    ),
    operation_gap(
        "remove",
        "slide",
        "referenced-page",
        "M6 ODP reference rewriting",
        "a slide (or a renamed slide) whose name links, custom shows or other markup reference is refused rather than leaving a dangling reference",
    ),
    operation_gap(
        "remove",
        "shape",
        "referenced-object",
        "M6 ODP reference rewriting",
        "an object whose draw:id or xml:id animations or other markup reference is refused; notes-page objects are not removed",
    ),
    operation_gap(
        "get",
        "metadata",
        "path-read",
        "File::metadata(); /metadata node reads with the shared metadata vocabulary",
        "`set(\"/metadata\", ..)` writes meta.xml; the path is set-only, and reads go through File::metadata() and the root node's properties",
    ),
    operation_gap(
        "set",
        "presentation",
        "extended-metadata",
        "0288 ODF user-defined and statistics metadata",
        "/metadata writes title, subject, description, author, lastModifiedBy, language and keywords through the ODF metadata family; dates, statistics and custom.<name> are refused",
    ),
];

const _: () = {
    let mut index = 0;
    while index < WRITE_GAPS.len() {
        assert!(crate::node::spec::gap_valid(&WRITE_GAPS[index]));
        index += 1;
    }
};
