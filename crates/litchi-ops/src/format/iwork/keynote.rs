//! Apple Keynote presentations (`.key`): the slides with their layout, the
//! title and body placeholders, speaker notes, the drawables each slide owns
//! (text boxes and shapes, pictures, lines, movies, audio, tables, charts
//! and groups) with their frames, and the paragraphs and character runs of
//! every text, read from `litchi_keynote::Package` (records 0296 and 0298).
//!
//! The vocabulary is the presentation family's, as PPTX and ODP read it:
//! `slide`, `placeholder`, `notes`, `shape` (text boxes and shapes),
//! `picture`, `connector`, `video`, `audio`, `table`/`tr`/`tc`, `chart`,
//! `group`, `paragraph` and `run`, with the frame properties `x`, `y`,
//! `width`, `height`, `rotation` and `zorder` and the run properties `bold`,
//! `italic`, `size`, `color` and `font.latin`.
//!
//! A text's paragraphs are the lines of its storage text: Keynote separates
//! paragraphs with a newline. Character runs are split at paragraph
//! boundaries and join the tree the first time a read reaches a paragraph's
//! children (reading them decodes each drawable's character-style table).

use litchi_core::FileFormat;
use litchi_keynote::slide::drawable::{Drawable, Kind as DrawableKind, TextRun};
use litchi_keynote::slide::placeholder::Kind as Role;
use litchi_keynote::{Package, SlideSelector, TableSelector};

use super::tree::{ROOT, count, text};
use super::write::leaf_refusal;
use super::{Flavor, LeafError, TreeBuilder, clearable_writable, element, writable};
use crate::Result;
use crate::error::Code;
use crate::node::Value;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::value::{Color, ColorContext, Length, LengthSpec, ValueType};

pub(crate) mod coalesce;
mod source;
mod write;

pub(crate) use coalesce::Run;
pub(crate) use write::{add, move_to, remove, set, swap};

kinds! {
    /// The element kinds of a Keynote presentation.
    Kind: &crate::path::vocabulary::PRESENTATION, root Presentation {
        Presentation => "presentation",
        Slide => "slide",
        Placeholder => "placeholder",
        Notes => "notes",
        Paragraph => "paragraph",
        Run => "run",
        Shape => "shape",
        Picture => "picture",
        Connector => "connector",
        Video => "video",
        Audio => "audio",
        Table => "table",
        Row => "tr",
        Cell => "tc",
        Chart => "chart",
        Group => "group",
    }
}

const POSITION: ValueType = ValueType::Length(LengthSpec::POSITION);
const EXTENT: ValueType = ValueType::Length(LengthSpec::EXTENT);
const FONT_SIZE: ValueType = ValueType::Length(LengthSpec::FONT_SIZE);
const COLOR: ValueType = ValueType::Color(ColorContext::Drawing);

/// A property `set` writes and `add` accepts.
const fn addable(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = writable(name, ty);
    spec.access.add = true;
    spec
}

/// A property only `add` accepts, read back under no key.
const fn add_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    spec.access.add = true;
    spec
}

/// A readable property `add` also accepts.
const fn readable_add(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.add = true;
    spec
}

const PRESENTATION_PROPS: &[PropSpec] = &[PropSpec::new("slideCount", ValueType::Int)];

pub(crate) const SLIDE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
    writable("hidden", ValueType::Bool),
    readable_add("layout", ValueType::Text),
    add_only("from", ValueType::Text),
];

pub(crate) const PLACEHOLDER_PROPS: &[PropSpec] = &[
    PropSpec::new("type", ValueType::Text),
    writable("text", ValueType::Text),
    writable("x", POSITION),
    writable("y", POSITION),
    writable("width", EXTENT),
    writable("height", EXTENT),
    writable("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
    writable("hidden", ValueType::Bool),
];

const NOTES_PROPS: &[PropSpec] = &[writable("text", ValueType::Text)];

const PARAGRAPH_PROPS: &[PropSpec] = &[addable("text", ValueType::Text)];

const RUN_PROPS: &[PropSpec] = &[
    writable("bold", ValueType::Bool),
    writable("italic", ValueType::Bool),
    writable("underline", ValueType::Bool),
    writable("size", FONT_SIZE),
    writable("color", COLOR),
    writable("font.latin", ValueType::Text),
];

const SHAPE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    addable("text", ValueType::Text),
    addable("x", POSITION),
    addable("y", POSITION),
    addable("width", EXTENT),
    addable("height", EXTENT),
    addable("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
    add_only("preset", ValueType::Text),
];

const PICTURE_PROPS: &[PropSpec] = &[
    readable_add("name", ValueType::Text),
    addable("x", POSITION),
    addable("y", POSITION),
    addable("width", EXTENT),
    addable("height", EXTENT),
    addable("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
    add_only("src", ValueType::Text),
];

/// The frame of a drawable no write reaches: lines, movies, audio, charts
/// and groups.
const FRAME_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("x", POSITION),
    PropSpec::new("y", POSITION),
    PropSpec::new("width", EXTENT),
    PropSpec::new("height", EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
];

const CHART_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("x", POSITION),
    PropSpec::new("y", POSITION),
    PropSpec::new("width", EXTENT),
    PropSpec::new("height", EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
    clearable_writable("font.size", FONT_SIZE),
];

const TABLE_PROPS: &[PropSpec] = &[
    readable_add("name", ValueType::Text),
    readable_add("rows", ValueType::Int),
    readable_add("cols", ValueType::Int),
    readable_add("x", POSITION),
    readable_add("y", POSITION),
    readable_add("width", EXTENT),
    readable_add("height", EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
    add_only("data", ValueType::Text),
];

const ROW_PROPS: &[PropSpec] = &[];

const CELL_PROPS: &[PropSpec] = &[
    writable("text", ValueType::Text),
    clearable_writable("comment", ValueType::Text),
    PropSpec::new("type", ValueType::Text),
    {
        let mut spec = PropSpec::new("value", ValueType::Scalar);
        spec.access = crate::node::spec::Access {
            get: false,
            set: true,
            add: false,
        };
        spec
    },
];

static GAPS: [GapSpec; 0] = [];

/// A conditional operation the Keynote leaf does not offer yet.
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

static OPERATION_GAPS: [GapSpec; 8] = [
    operation_gap(
        "add",
        "run",
        "a new character run",
        "IWORK_API_REQUESTS.md request 10 (Keynote character-style writes)",
        "a run is a styled span of a paragraph; add a paragraph or set its text",
    ),
    operation_gap(
        "set",
        "tc",
        "cells in a formula graph",
        "record 0309 follow-ups",
        "slide-table cells take text, numbers, booleans and clears; a formula cell, or one a formula or chart reads, is refused so no calculation cache goes stale",
    ),
    operation_gap(
        "add",
        "table",
        "a deck without a slide table",
        "record 0329 follow-up: a built-in table template",
        "a new table copies the structure (styles, header counts, storage layout) of the deck's first slide table; a deck without one has none to copy",
    ),
    operation_gap(
        "remove",
        "table",
        "a slide table, chart or group",
        "record 0298 follow-up: structural removal of tables, charts and groups",
        "the leaf's drawable removal refuses tables, charts and groups",
    ),
    operation_gap(
        "add",
        "video",
        "a new movie or audio clip",
        "record 0309 follow-up: wire slide::movie and slide::audio creation",
        "the leaf's movie and audio creation families need the clip's poster, duration and placement, which `add` does not take yet",
    ),
    operation_gap(
        "set",
        "video",
        "the frame of a movie, audio clip, line, table, chart or group",
        "record 0298 follow-up: wire slide::movie::geometry; the drawable family refuses the others",
        "the leaf's drawable geometry refuses tables, charts, groups, lines and movies",
    ),
    operation_gap(
        "remove",
        "placeholder",
        "slide-number and object placeholders",
        "record 0309 follow-ups",
        "`remove` hides a title or body placeholder on its slide (and `hidden` = false shows it again); slide-number and object placeholders have no unified element",
    ),
    operation_gap(
        "add",
        "placeholder",
        "a new placeholder",
        "record 0309 follow-ups",
        "a slide's placeholders come from its layout; show a hidden one with `hidden` = false, or add a slide from a layout",
    ),
];

/// The Keynote spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Presentation,
    elements: &[
        element(
            Kind::Presentation,
            "presentation",
            &["presentation"],
            &[],
            &[Kind::Slide],
            &["/"],
            PRESENTATION_PROPS,
        ),
        element(
            Kind::Slide,
            "slide",
            &["slide"],
            &[Kind::Presentation],
            &[
                Kind::Placeholder,
                Kind::Notes,
                Kind::Shape,
                Kind::Picture,
                Kind::Connector,
                Kind::Video,
                Kind::Audio,
                Kind::Table,
                Kind::Chart,
                Kind::Group,
            ],
            &["/slide[N]"],
            SLIDE_PROPS,
        ),
        element(
            Kind::Placeholder,
            "placeholder",
            &["ph", "placeholder"],
            &[Kind::Slide],
            &[Kind::Paragraph],
            &[
                "/slide[N]/placeholder[title]",
                "/slide[N]/placeholder[body]",
            ],
            PLACEHOLDER_PROPS,
        ),
        element(
            Kind::Notes,
            "notes",
            &["notes"],
            &[Kind::Slide],
            &[Kind::Paragraph],
            &["/slide[N]/notes"],
            NOTES_PROPS,
        ),
        element(
            Kind::Paragraph,
            "paragraph",
            &["p", "paragraph"],
            &[Kind::Placeholder, Kind::Notes, Kind::Shape],
            &[Kind::Run],
            &[
                "/slide[N]/placeholder[title]/paragraph[N]",
                "/slide[N]/notes/paragraph[N]",
                "/slide[N]/shape[N]/paragraph[N]",
            ],
            PARAGRAPH_PROPS,
        ),
        element(
            Kind::Run,
            "run",
            &["r", "run"],
            &[Kind::Paragraph],
            &[],
            &["/slide[N]/shape[N]/paragraph[N]/run[N]"],
            RUN_PROPS,
        ),
        element(
            Kind::Shape,
            "shape",
            &["shape", "textbox"],
            &[Kind::Slide],
            &[Kind::Paragraph],
            &["/slide[N]/shape[N]"],
            SHAPE_PROPS,
        ),
        element(
            Kind::Picture,
            "picture",
            &["image", "img", "pic", "picture"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/picture[N]"],
            PICTURE_PROPS,
        ),
        element(
            Kind::Connector,
            "connector",
            &["connection", "connector"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/connector[N]"],
            FRAME_PROPS,
        ),
        element(
            Kind::Video,
            "video",
            &["video"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/video[N]"],
            FRAME_PROPS,
        ),
        element(
            Kind::Audio,
            "audio",
            &["audio"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/audio[N]"],
            FRAME_PROPS,
        ),
        element(
            Kind::Table,
            "table",
            &["table", "tbl"],
            &[Kind::Slide],
            &[Kind::Row],
            &["/slide[N]/table[N]"],
            TABLE_PROPS,
        ),
        element(
            Kind::Row,
            "tr",
            &["row", "tr"],
            &[Kind::Table],
            &[Kind::Cell],
            &["/slide[N]/table[N]/tr[N]"],
            ROW_PROPS,
        ),
        element(
            Kind::Cell,
            "tc",
            &["cell", "tc"],
            &[Kind::Row],
            &[],
            &["/slide[N]/table[N]/tr[N]/tc[N]"],
            CELL_PROPS,
        ),
        element(
            Kind::Chart,
            "chart",
            &["chart"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/chart[N]"],
            CHART_PROPS,
        ),
        element(
            Kind::Group,
            "group",
            &["group"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/group[N]"],
            FRAME_PROPS,
        ),
    ],
    gaps: &GAPS,
};

crate::node::spec::check_format_spec!(SPEC);

/// The Keynote flavor.
pub(crate) struct Keynote;

impl Flavor for Keynote {
    type Kind = Kind;
    const FORMAT: FileFormat = FileFormat::Keynote;
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Presentation];
    const SINGLETONS: &'static [Kind] = &[Kind::Notes];
}

/// The proof of one Keynote leaf transaction.
#[allow(
    clippy::large_enum_variant,
    reason = "a write holds its proofs once; boxing would add an allocation per commit"
)]
pub(crate) enum Patch {
    Text(litchi_keynote::slide::text::Patch),
    Notes(litchi_keynote::slide::notes::Patch),
    Skip(litchi_keynote::edit::Patch),
    Order(litchi_keynote::slide::order::Patch),
    Deletion(
        #[allow(
            dead_code,
            reason = "retained as the leaf's reversible proof; a committed deletion is never a no-op"
        )]
        litchi_keynote::slide::delete::Patch,
    ),
    Drawables(litchi_keynote::slide::drawable::Patch),
    ChartFontSize(litchi_keynote::chart::font::size::Patch),
    TableComments(litchi_keynote::slide::table::comments::Patch),
    Creation(litchi_keynote::slide::creation::Patch),
    Media(litchi_keynote::slide::media::lifecycle::Patch),
    Placeholder(litchi_keynote::slide::placeholder::Patch),
}

impl Patch {
    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Text(patch) => patch.is_noop(),
            Self::Notes(patch) => patch.is_noop(),
            Self::Skip(patch) => patch.is_noop(),
            Self::Order(patch) => patch.is_noop(),
            // A committed deletion always removes one slide.
            Self::Deletion(_) => false,
            Self::Drawables(patch) => patch.is_noop(),
            Self::ChartFontSize(patch) => patch.is_noop(),
            Self::TableComments(patch) => patch.is_noop(),
            Self::Creation(patch) => patch.is_noop(),
            Self::Media(patch) => patch.is_noop(),
            Self::Placeholder(patch) => patch.is_noop(),
        }
    }
}

/// The elements `set` writes, `remove` deletes and `move` reorders.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &[
        "slide",
        "placeholder",
        "notes",
        "paragraph",
        "run",
        "shape",
        "picture",
        "tc",
        "chart",
    ],
    addable: &["slide", "paragraph", "shape", "picture", "table"],
    removable: &[
        "slide",
        "paragraph",
        "shape",
        "picture",
        "connector",
        "video",
        "audio",
        "placeholder",
    ],
    movable: &[
        "slide",
        "paragraph",
        "placeholder",
        "shape",
        "picture",
        "connector",
        "video",
        "audio",
        "table",
        "chart",
        "group",
    ],
    gaps: &OPERATION_GAPS,
};

/// What a text-bearing item's `leaf[1]` records: the title or body
/// placeholder, the speaker notes, or a drawable's own storage
/// (`leaf[3]`). A drawable's `leaf[1]` is [`DRAWABLE`]; `leaf[2]` is a
/// paragraph's line, `leaf[3]` the drawable's position in the slide's
/// drawable list ([`NONE`] for a placeholder the slide lists no frame of).
const TITLE: usize = 0;
const BODY: usize = 1;
const NOTES: usize = 2;
const DRAWABLE: usize = 3;
const NONE: usize = u32::MAX as usize;

/// Points to EMU, rounded.
fn emu(points: f32) -> Length {
    #[allow(
        clippy::cast_possible_truncation,
        reason = "a frame coordinate in points is far inside the i64 range"
    )]
    Length::from_emu((f64::from(points) * 12_700.0).round() as i64)
}

/// Set a drawable's frame and z-order properties on an item.
fn frame<'b>(
    draft: super::tree::Draft<'b, Kind>,
    drawable: Option<&Drawable>,
) -> super::tree::Draft<'b, Kind> {
    let geometry = drawable.and_then(Drawable::geometry);
    let zorder = drawable
        .and_then(Drawable::z_position)
        .and_then(|position| count(position.get()));
    draft
        .prop("x", geometry.map(|frame| Value::Length(emu(frame.x))))
        .prop("y", geometry.map(|frame| Value::Length(emu(frame.y))))
        .prop(
            "width",
            geometry.map(|frame| Value::Length(emu(frame.width))),
        )
        .prop(
            "height",
            geometry.map(|frame| Value::Length(emu(frame.height))),
        )
        .prop(
            "rotation",
            geometry
                .and_then(|frame| frame.rotation)
                .and_then(|degrees| Value::number(f64::from(degrees))),
        )
        .prop("zorder", zorder)
}

/// The element kind of a drawable of a slide, or `None` for the
/// placeholders (title and body read through the slide-text projection;
/// slide-number and object placeholders have no unified element) and kinds
/// without one.
const fn element_of(kind: DrawableKind) -> Option<Kind> {
    Some(match kind {
        DrawableKind::TextBox | DrawableKind::Shape => Kind::Shape,
        DrawableKind::Image => Kind::Picture,
        DrawableKind::Line => Kind::Connector,
        DrawableKind::Movie => Kind::Video,
        DrawableKind::Audio => Kind::Audio,
        DrawableKind::Table => Kind::Table,
        DrawableKind::Chart => Kind::Chart,
        DrawableKind::Group => Kind::Group,
        _ => return None,
    })
}

/// The paragraphs of a text whose runs join the tree later.
struct Deferred {
    /// The paragraph items, in text order.
    paragraphs: Vec<u32>,
    slide: usize,
    /// The drawable holding the text, or [`NONE`] for the speaker notes
    /// (record 0329).
    drawable: usize,
}

/// Push the paragraphs of `value` under `parent`.
fn paragraphs(
    tree: &mut TreeBuilder<Kind>,
    parent: u32,
    value: &str,
    [slide, container, drawable]: [usize; 3],
) -> Vec<u32> {
    value
        .split('\n')
        .enumerate()
        .map(|(line, paragraph)| {
            tree.push(parent, Kind::Paragraph)
                .prop("text", text(Some(paragraph)))
                .text(paragraph)
                .leaf([slide, container, line, drawable])
                .done()
        })
        .collect()
}

/// Push the character runs of one text under its paragraphs, split at the
/// paragraph boundaries.
fn runs(tree: &mut TreeBuilder<Kind>, deferred: &Deferred, runs: &[TextRun]) {
    let mut start = 0_usize;
    let mut runs = runs.iter().peekable();
    for (line, paragraph) in deferred.paragraphs.iter().enumerate() {
        let length = tree.item_text(*paragraph).chars().count();
        let end = start + length;
        while let Some(run) = runs.peek() {
            let span = run.span();
            if span.start() >= end {
                break;
            }
            let (from, to) = (span.start().max(start), span.end().min(end));
            if from < to {
                let piece = run
                    .text()
                    .chars()
                    .skip(from - span.start())
                    .take(to - from)
                    .collect::<String>();
                let style = run.style();
                tree.push(*paragraph, Kind::Run)
                    .prop("bold", style.bold.map(Value::Bool))
                    .prop("italic", style.italic.map(Value::Bool))
                    .prop("underline", style.underline.map(Value::Bool))
                    .prop(
                        "size",
                        style.font_size.map(|points| Value::Length(emu(points))),
                    )
                    .prop(
                        "color",
                        style
                            .color
                            .map(|[red, green, blue]| Value::Color(Color::rgb(red, green, blue))),
                    )
                    .prop("font.latin", text(style.font_name.as_deref()))
                    .text(piece)
                    .leaf([
                        deferred.slide,
                        if deferred.drawable == NONE {
                            NOTES
                        } else {
                            DRAWABLE
                        },
                        line,
                        deferred.drawable,
                    ])
                    .done();
            }
            if span.end() <= end {
                runs.next();
            } else {
                break;
            }
        }
        // The newline after the paragraph.
        start = end + 1;
    }
}

/// Build the eager tree of `package`: the presentation, whose slides are a
/// lazily materialized part ([`Slides`]); each slide's children are a
/// lazily built part ([`build_slide`]), built the first time a read
/// reaches them (record 0305).
pub(crate) fn build(package: &Package) -> Result<(super::Tree, Box<[super::Part]>)> {
    let read = |error| leaf_refusal(LeafError::Keynote(error), FileFormat::Keynote);
    let slides = package.slides().map_err(read)?;
    let mut tree = TreeBuilder::new(SPEC.elements, Keynote::SINGLETONS);
    tree.root()
        .prop("slideCount", count(slides.len()))
        .part(Slides::PART)
        .done();
    let mut parts = Vec::with_capacity(slides.len() + 1);
    parts.push(super::Part::Slides(Slides::new(slides.len())));
    for index in 0..slides.len() {
        parts.push(super::Part::Slide(super::part::Whole::new(
            index,
            Box::new([slide_step(index)]),
        )));
    }
    Ok((tree.build(), parts.into_boxed_slice()))
}

/// Build only one slide's chart structural items for the canonical chart path
/// fast route. The ordinary slide tree remains the source of all other reads;
/// unrelated slide inventories are never inspected.
pub(crate) fn build_charts(package: &Package, slide_index: usize) -> Result<super::Tree> {
    let read = |error| leaf_refusal(LeafError::Keynote(error), FileFormat::Keynote);
    let slides = package.slides().map_err(read)?;
    let mut tree = TreeBuilder::new(SPEC.elements, Keynote::SINGLETONS);
    tree.root().prop("slideCount", count(slides.len())).done();
    let Some(_) = slides.get(slide_index) else {
        return Ok(tree.build());
    };
    let slide = tree
        .push_step(
            ROOT,
            Kind::Slide,
            crate::node::PathStep::at("slide", litchi_core::Position::new(slide_index)),
        )
        .leaf([slide_index, 0, 0, 0])
        .done();
    let drawables = package
        .slide_drawable_list(SlideSelector::index(slide_index))
        .map_err(|error| leaf_refusal(LeafError::KeynoteDrawables(error), FileFormat::Keynote))?;
    let mut chart_position = 0;
    for (position, drawable) in drawables.iter().enumerate() {
        if drawable.kind() != DrawableKind::Chart {
            continue;
        }
        frame(
            tree.push(slide, Kind::Chart)
                .leaf([slide_index, DRAWABLE, chart_position, position]),
            Some(drawable),
        )
        .done();
        chart_position += 1;
    }
    Ok(tree.build())
}

/// The canonical step of slide `index`.
fn slide_step(index: usize) -> crate::node::PathStep<'static> {
    crate::node::PathStep::at(
        crate::path::navigate::ElementKind::name(Kind::Slide),
        litchi_core::Position::new(index),
    )
}

/// The slides of a presentation, each item materialized the first time a
/// read reaches it (record 0305): a slide's `layout` reads the theme's
/// layout catalog, decoded once, on the first slide item.
pub(crate) struct Slides {
    items: Box<[std::sync::OnceLock<Box<super::tree::Item>>]>,
    /// The layout each slide was made from; `None` when the layout reader
    /// refuses the theme or a slide record (`layout` is then unset; it
    /// never hides the slides).
    layouts: std::sync::OnceLock<Option<Box<[Option<litchi_keynote::slide::creation::Layout>]>>>,
}

impl Slides {
    /// The kind ordinal of every item.
    pub(crate) const KIND: u8 = Kind::Slide as u8;
    /// The slides' part: the first.
    const PART: u32 = 1;

    fn new(count: usize) -> Self {
        Self {
            items: (0..count).map(|_| std::sync::OnceLock::new()).collect(),
            layouts: std::sync::OnceLock::new(),
        }
    }

    /// The part holding the children of slide `slot`: the parts after the
    /// slides' own.
    pub(crate) const fn child_part(slot: u32) -> u32 {
        slot.saturating_add(Self::PART + 1)
    }

    pub(crate) fn len(&self) -> u32 {
        u32::try_from(self.items.len()).unwrap_or(u32::MAX)
    }

    #[cfg(test)]
    pub(crate) fn state(&self) -> String {
        let built = self
            .items
            .iter()
            .filter(|item| item.get().is_some())
            .count();
        let layouts = if self.layouts.get().is_some() {
            "+layouts"
        } else {
            ""
        };
        format!("slides {built}/{}{layouts}", self.items.len())
    }

    /// The item of slide `slot`, materialized on the first call.
    pub(crate) fn item<'f>(&'f self, file: &'f super::Iwork, slot: u32) -> &'f super::tree::Item {
        self.items[slot as usize].get_or_init(|| Box::new(self.slide(file, slot as usize)))
    }

    fn slide(&self, file: &super::Iwork, index: usize) -> super::tree::Item {
        let spec = &SPEC.elements[Kind::Slide as usize];
        let ordinal = u32::try_from(index).unwrap_or(u32::MAX);
        let mut item = super::tree::Item::detached(spec, Box::new([slide_step(index)]), ordinal);
        let Some(package) = file.keynote() else {
            return item;
        };
        let layouts = self
            .layouts
            .get_or_init(|| package.slide_layout_list().ok());
        let layout = layouts
            .as_ref()
            .and_then(|layouts| layouts.get(index))
            .and_then(Option::as_ref)
            .and_then(|layout| layout.name());
        // The handle's package is immutable: its slides decoded when the
        // eager tree was built.
        let slide = package.slides().ok().and_then(|slides| slides.get(index));
        let draft = item.draft::<Kind>();
        match slide {
            Some(slide) => draft
                .prop("name", text(slide.name()))
                .prop("title", text(slide.title()))
                .prop("hidden", Some(slide.is_skipped().into()))
                .prop("layout", text(layout))
                .label(slide.name()),
            None => draft,
        }
        .leaf([index, 0, 0, 0])
        .part(Self::child_part(ordinal))
        .done();
        item
    }
}

/// Build the children of slide `index` (at `path`): its placeholders,
/// notes and drawables with their paragraphs, and with `with_runs` the
/// character runs of every text, pushed last ([`TreeBuilder::defer`]):
/// reading them decodes every drawable's character-style table, so they
/// join the part the first time a read reaches a paragraph's children.
/// Every item of the part without runs keeps its index.
///
/// A table is a stand-in holding its frame (record 0312): reading a table
/// decodes it, so its name, size, rows and cells are built the first time
/// a read reaches them ([`build_table`]). Returns the tree and, for each
/// stand-in in order, the table's selector (its position among the
/// slide's tables in z-order) and the stand-in's index.
pub(crate) fn build_slide(
    package: &Package,
    index: usize,
    path: &[crate::node::PathStep<'static>],
    with_runs: bool,
) -> (super::Tree, Vec<(usize, u32)>) {
    let mut tree = TreeBuilder::under(SPEC.elements, Keynote::SINGLETONS, Kind::Slide, path);
    let mut selectors = Vec::new();
    let Some(slide) = package.slides().ok().and_then(|slides| slides.get(index)) else {
        return (tree.build(), selectors);
    };
    let item = ROOT;
    let mut deferred = Vec::new();
    let selector = SlideSelector::index(index);
    // A slide the drawable inventory refuses keeps its placeholders and
    // notes; it lists no drawables.
    let drawables = package.slide_drawable_list(selector).ok();
    let drawables = drawables.as_deref().unwrap_or_default();
    // The placeholders the leaf's slide-text transaction addresses; a
    // slide without one has no such child.
    let roles = [
        (TITLE, Role::Title, "title", DrawableKind::TitlePlaceholder),
        (BODY, Role::Body, "body", DrawableKind::BodyPlaceholder),
    ];
    for (slot, role, name, kind) in roles {
        let Ok(Some(value)) = package.slide_text(selector, role) else {
            continue;
        };
        let position = drawables
            .iter()
            .position(|drawable| drawable.kind() == kind);
        let drawable = position.and_then(|position| drawables.get(position));
        // A hidden placeholder keeps its text on its slide (record 0312).
        let hidden = package
            .slide_placeholder_visibility(selector, role)
            .ok()
            .flatten()
            .map(|state| {
                Value::Bool(matches!(
                    state,
                    litchi_keynote::slide::placeholder::State::Hidden
                ))
            });
        let placeholder = frame(
            tree.push(item, Kind::Placeholder)
                .prop("type", text(Some(name)))
                .prop("text", text(Some(&value)))
                .prop("hidden", hidden),
            drawable,
        )
        .label(Some(name))
        .text(value.clone())
        .leaf([index, slot, 0, position.unwrap_or(NONE)])
        .done();
        let lines = paragraphs(&mut tree, placeholder, &value, [index, slot, NONE]);
        if let Some(position) = position {
            deferred.push(Deferred {
                paragraphs: lines,
                slide: index,
                drawable: position,
            });
        }
    }
    if let Some(notes) = slide.notes() {
        let item = tree
            .push(item, Kind::Notes)
            .prop("text", text(Some(notes)))
            .text(notes)
            .leaf([index, NOTES, 0, NONE])
            .done();
        let lines = paragraphs(&mut tree, item, notes, [index, NOTES, NONE]);
        deferred.push(Deferred {
            paragraphs: lines,
            slide: index,
            drawable: NONE,
        });
    }
    let mut tables = drawables
        .iter()
        .enumerate()
        .filter(|(_, drawable)| drawable.kind() == DrawableKind::Table)
        .map(|(position, drawable)| (drawable.z_position(), position))
        .collect::<Vec<_>>();
    // A table selector addresses the slide's tables in z-order.
    tables.sort_unstable();
    for (position, drawable) in drawables.iter().enumerate() {
        let Some(kind) = element_of(drawable.kind()) else {
            continue;
        };
        let chart_position = (kind == Kind::Chart).then(|| {
            drawables[..position]
                .iter()
                .filter(|candidate| candidate.kind() == DrawableKind::Chart)
                .count()
        });
        let draft = tree
            .push(item, kind)
            .prop("name", text(drawable.name()))
            .label(drawable.name())
            .leaf([index, DRAWABLE, chart_position.unwrap_or(0), position]);
        let draft = match kind {
            Kind::Shape => {
                let value = drawable.text().unwrap_or_default();
                draft.prop("text", text(Some(value))).text(value)
            },
            _ => draft,
        };
        let draft = match kind {
            // The stand-in's part names its contents among the slide's
            // tables; a slide with more tables than a part addresses keeps
            // the later ones' frames without rows.
            Kind::Table if selectors.len() < super::part::SLIDE_TABLES => {
                let ordinal = tables
                    .iter()
                    .position(|(_, table)| *table == position)
                    .unwrap_or(0);
                selectors.push((ordinal, 0));
                draft.part(u32::try_from(selectors.len()).unwrap_or(u32::MAX))
            },
            _ => draft,
        };
        let drawable_item = frame(draft, Some(drawable)).done();
        if kind == Kind::Table
            && let Some(last) = selectors.last_mut()
            && last.1 == 0
        {
            last.1 = drawable_item;
        }
        if kind == Kind::Shape {
            let value = drawable.text().unwrap_or_default();
            let lines = paragraphs(&mut tree, drawable_item, value, [index, DRAWABLE, position]);
            if drawable.text().is_some() {
                deferred.push(Deferred {
                    paragraphs: lines,
                    slide: index,
                    drawable: position,
                });
            }
        }
    }
    if with_runs {
        tree.defer();
        for text in &deferred {
            let read = if text.drawable == NONE {
                package.slide_notes_text_runs(SlideSelector::index(text.slide))
            } else {
                package.slide_drawable_text_runs(
                    SlideSelector::index(text.slide),
                    litchi_keynote::DrawableSelector::index(text.drawable),
                )
            };
            let Ok(read) = read else {
                continue;
            };
            runs(&mut tree, text, &read);
        }
    }
    (tree.build(), selectors)
}

/// Build the contents of the slide table whose stand-in is `stand_in` and
/// whose selector among its slide's tables is `ordinal`: its name, size,
/// rows and cells (record 0312). A table the leaf's cell reader refuses,
/// or one with more cells than a slide's part addresses, keeps its frame
/// without rows.
pub(crate) fn build_table(
    package: &Package,
    stand_in: &super::tree::Item,
    ordinal: usize,
) -> super::Tree {
    let mut tree = TreeBuilder::from_stand_in(SPEC.elements, Keynote::SINGLETONS, stand_in);
    let [slide, _, _, drawable] = stand_in.leaf().map(|value| value as usize);
    let Ok(read) =
        package.slide_table_cells(SlideSelector::index(slide), TableSelector::index(ordinal))
    else {
        return tree.build();
    };
    let table = read.table();
    let comments = read
        .iter_comments()
        .map(|comment| {
            (
                comment.position().row() as usize,
                comment.position().column() as usize,
                comment.comment().text().to_owned(),
            )
        })
        .map(|(row, column, text)| ((row, column), text))
        .collect::<std::collections::BTreeMap<_, _>>();
    let dimensions = table.dimensions();
    let (rows, columns) = (dimensions.rows(), dimensions.columns());
    tree.root()
        .prop("name", text(Some(table.name())))
        .prop("rows", count(rows as usize))
        .prop("cols", count(columns as usize))
        .label(Some(table.name()))
        .done();
    let items = (rows as usize).saturating_mul(1 + columns as usize);
    if items >= super::part::TABLE_ITEMS {
        return tree.build();
    }
    for row in 0..rows {
        let row_item = tree
            .push(ROOT, Kind::Row)
            .leaf([slide, DRAWABLE, row as usize, drawable])
            .done();
        for column in 0..columns {
            let value = table.get(litchi_keynote::slide::table::cells::At::new(row, column));
            let shown = value.map(ToString::to_string);
            tree.push(row_item, Kind::Cell)
                .prop("text", text(shown.as_deref()))
                .prop("type", text(value.map(value_type)))
                .prop(
                    "comment",
                    text(
                        comments
                            .get(&(row as usize, column as usize))
                            .map(String::as_str),
                    ),
                )
                .text(shown.unwrap_or_default())
                .leaf([slide, DRAWABLE, row as usize, drawable])
                .done();
        }
    }
    debug_assert!(tree.len() <= super::part::TABLE_ITEMS);
    tree.build()
}

/// The type token of a table cell value.
fn value_type(value: &litchi_keynote::slide::table::cells::Value) -> &'static str {
    use litchi_keynote::slide::table::cells::Value as V;
    match value {
        V::Text(_) => "text",
        V::Number(_) => "number",
        V::Boolean(_) => "boolean",
        V::Date(_) => "date",
        V::Duration(_) => "duration",
        V::Formula(_) => "formula",
        V::Error(_) => "error",
        V::Empty => "empty",
    }
}
