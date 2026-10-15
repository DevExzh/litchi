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

use litchi_core::{FileFormat, Position};
use litchi_keynote::slide::drawable::{Drawable, Kind as DrawableKind, TextRun};
use litchi_keynote::slide::placeholder::Kind as Role;
use litchi_keynote::{Package, SlideSelector, TableSelector};

use super::tree::{ROOT, count, text};
use super::write::{
    Chain, Commit, LeafPatch, bound, destination, element_refusal, finish, key_refusal,
    leaf_refusal, node_at, refusal, resolve,
};
use super::{Flavor, Iwork, LeafError, Owner, TreeBuilder, element, writable};
use crate::Result;
use crate::error::Code;
use crate::node::Value;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::path::PathArg;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::{Color, ColorContext, Length, LengthSpec, ValueType};

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

const PRESENTATION_PROPS: &[PropSpec] = &[PropSpec::new("slideCount", ValueType::Int)];

pub(crate) const SLIDE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
    writable("hidden", ValueType::Bool),
    PropSpec::new("layout", ValueType::Text),
];

pub(crate) const PLACEHOLDER_PROPS: &[PropSpec] = &[
    PropSpec::new("type", ValueType::Text),
    writable("text", ValueType::Text),
    PropSpec::new("x", POSITION),
    PropSpec::new("y", POSITION),
    PropSpec::new("width", EXTENT),
    PropSpec::new("height", EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
];

const NOTES_PROPS: &[PropSpec] = &[PropSpec::new("text", ValueType::Text)];

const PARAGRAPH_PROPS: &[PropSpec] = &[PropSpec::new("text", ValueType::Text)];

const RUN_PROPS: &[PropSpec] = &[
    PropSpec::new("bold", ValueType::Bool),
    PropSpec::new("italic", ValueType::Bool),
    PropSpec::new("underline", ValueType::Bool),
    PropSpec::new("size", FONT_SIZE),
    PropSpec::new("color", COLOR),
    PropSpec::new("font.latin", ValueType::Text),
];

const SHAPE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("text", ValueType::Text),
    PropSpec::new("x", POSITION),
    PropSpec::new("y", POSITION),
    PropSpec::new("width", EXTENT),
    PropSpec::new("height", EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
];

const PICTURE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("x", POSITION),
    PropSpec::new("y", POSITION),
    PropSpec::new("width", EXTENT),
    PropSpec::new("height", EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
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

const TABLE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("rows", ValueType::Int),
    PropSpec::new("cols", ValueType::Int),
    PropSpec::new("x", POSITION),
    PropSpec::new("y", POSITION),
    PropSpec::new("width", EXTENT),
    PropSpec::new("height", EXTENT),
    PropSpec::new("rotation", ValueType::Number),
    PropSpec::new("zorder", ValueType::Int),
];

const ROW_PROPS: &[PropSpec] = &[];

const CELL_PROPS: &[PropSpec] = &[
    PropSpec::new("text", ValueType::Text),
    PropSpec::new("type", ValueType::Text),
];

static GAPS: [GapSpec; 1] = [GapSpec {
    subject: Subject::Operation {
        verb: "set",
        element: "notes",
        context: "speaker notes text",
    },
    code: Code::UnsupportedCapability,
    follow_up: "0298 Keynote notes writes",
    reason: "speaker notes are read-only here; the leaf's slide-notes transaction is not wired yet",
}];

static OPERATION_GAPS: [GapSpec; 2] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "slide",
            context: "a new slide",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 8",
        reason: "the Keynote leaf has no slide creation transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "presentation",
            context: "a presentation from scratch",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 1",
        reason: "the Keynote leaf has no from-scratch package builder",
    },
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
            &["table"],
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
            FRAME_PROPS,
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

/// The elements `set` writes, `remove` deletes and `move` reorders.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["placeholder", "slide"],
    addable: &[],
    removable: &["slide"],
    movable: &["slide"],
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
                    .leaf([deferred.slide, DRAWABLE, line, deferred.drawable])
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

/// Build the tree of `package` without character runs: reading them decodes
/// every drawable's character-style table, so they join the tree the first
/// time a read reaches a paragraph's children ([`build_full`]).
pub(crate) fn build(package: &Package) -> Result<super::Tree> {
    build_with(package, false)
}

/// Build the complete tree of `package`, character runs included. Every
/// item of [`build`]'s tree keeps its index: the runs are pushed last.
pub(crate) fn build_full(package: &Package) -> Result<super::Tree> {
    build_with(package, true)
}

fn build_with(package: &Package, with_runs: bool) -> Result<super::Tree> {
    let read = |error| leaf_refusal(LeafError::Keynote(error), FileFormat::Keynote);
    let slides = package.slides().map_err(read)?;
    // A theme or slide record the layout reader refuses leaves `layout`
    // unset; it never hides the slides.
    let layouts = package.slide_layout_list().ok();
    let mut tree = TreeBuilder::new(SPEC.elements, Keynote::SINGLETONS);
    tree.root().prop("slideCount", count(slides.len())).done();
    let mut deferred = Vec::new();
    for (index, slide) in slides.iter().enumerate() {
        let layout = layouts
            .as_ref()
            .and_then(|layouts| layouts.get(index))
            .and_then(Option::as_ref)
            .and_then(|layout| layout.name());
        let item = tree
            .push(ROOT, Kind::Slide)
            .prop("name", text(slide.name()))
            .prop("title", text(slide.title()))
            .prop("hidden", Some(slide.is_skipped().into()))
            .prop("layout", text(layout))
            .label(slide.name())
            .leaf([index, 0, 0, 0])
            .done();
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
            let placeholder = frame(
                tree.push(item, Kind::Placeholder)
                    .prop("type", text(Some(name)))
                    .prop("text", text(Some(&value))),
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
            paragraphs(&mut tree, item, notes, [index, NOTES, NONE]);
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
            let draft = tree
                .push(item, kind)
                .prop("name", text(drawable.name()))
                .label(drawable.name())
                .leaf([index, DRAWABLE, 0, position]);
            let draft = match kind {
                Kind::Shape => {
                    let value = drawable.text().unwrap_or_default();
                    draft.prop("text", text(Some(value))).text(value)
                },
                _ => draft,
            };
            let drawable_item = frame(draft, Some(drawable)).done();
            match kind {
                Kind::Shape => {
                    let value = drawable.text().unwrap_or_default();
                    let lines =
                        paragraphs(&mut tree, drawable_item, value, [index, DRAWABLE, position]);
                    if drawable.text().is_some() {
                        deferred.push(Deferred {
                            paragraphs: lines,
                            slide: index,
                            drawable: position,
                        });
                    }
                },
                Kind::Table => {
                    let ordinal = tables
                        .iter()
                        .position(|(_, table)| *table == position)
                        .unwrap_or(0);
                    table(&mut tree, package, drawable_item, index, ordinal, position);
                },
                _ => {},
            }
        }
    }
    if with_runs {
        for text in &deferred {
            let Ok(read) = package.slide_drawable_text_runs(
                SlideSelector::index(text.slide),
                litchi_keynote::DrawableSelector::index(text.drawable),
            ) else {
                continue;
            };
            runs(&mut tree, text, &read);
        }
    }
    Ok(tree.build())
}

/// Read one table's name, size and cells. A table the leaf's cell reader
/// refuses keeps its frame without rows.
fn table(
    tree: &mut TreeBuilder<Kind>,
    package: &Package,
    item: u32,
    slide: usize,
    ordinal: usize,
    drawable: usize,
) {
    let Ok(read) =
        package.slide_table_cells(SlideSelector::index(slide), TableSelector::index(ordinal))
    else {
        return;
    };
    let table = read.table();
    let dimensions = table.dimensions();
    let (rows, columns) = (dimensions.rows(), dimensions.columns());
    tree.draft(item)
        .prop("name", text(Some(table.name())))
        .prop("rows", count(rows as usize))
        .prop("cols", count(columns as usize))
        .label(Some(table.name()))
        .done();
    for row in 0..rows {
        let row_item = tree
            .push(item, Kind::Row)
            .leaf([slide, DRAWABLE, row as usize, drawable])
            .done();
        for column in 0..columns {
            let value = table.get(litchi_keynote::slide::table::cells::CellPosition::new(
                row, column,
            ));
            let shown = value.map(ToString::to_string);
            tree.push(row_item, Kind::Cell)
                .prop("text", text(shown.as_deref()))
                .prop("type", text(value.map(value_type)))
                .text(shown.unwrap_or_default())
                .leaf([slide, DRAWABLE, row as usize, drawable])
                .done();
        }
    }
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

/// Set slide visibility or placeholder text on every target.
pub(crate) fn set(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Keynote>(file, target, limits.max_expanded_targets())?;
    let mut chain = Chain::new(package.clone());
    let mut operations = 0_usize;
    for node in &nodes {
        for (key, value) in props.iter() {
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            let key = key.as_str();
            let [slide, slot, _, _] = node.item().leaf().map(|value| value as usize);
            let selector = SlideSelector::index(slide);
            match (node.item().type_name(), key) {
                ("slide", "hidden") => {
                    let hidden = super::write::boolean(value, "hidden")?;
                    let error =
                        |error| leaf_refusal(LeafError::KeynoteEdit(error), FileFormat::Keynote);
                    let mut edit = chain.package.edit();
                    edit.set_slide_skipped(selector, hidden).map_err(error)?;
                    let commit = edit.commit().map_err(error)?;
                    let patch = commit.patch().clone();
                    let changed = !patch.is_noop();
                    chain.push(
                        commit.into_package(),
                        LeafPatch::KeynoteSkip(patch),
                        changed,
                    );
                },
                ("placeholder", "text") => {
                    let value = super::write::text(value, "text")?.unwrap_or_default();
                    let role = if slot == TITLE {
                        Role::Title
                    } else {
                        Role::Body
                    };
                    let error =
                        |error| leaf_refusal(LeafError::KeynoteText(error), FileFormat::Keynote);
                    let mut edit = chain
                        .package
                        .edit_slide_text(selector, role)
                        .map_err(error)?;
                    edit.set(&value).map_err(error)?;
                    let commit = edit.commit().map_err(error)?;
                    let patch = commit.patch().clone();
                    let changed = !patch.is_noop();
                    chain.push(
                        commit.into_package(),
                        LeafPatch::KeynoteText(patch),
                        changed,
                    );
                },
                ("slide" | "placeholder", _) => return Err(key_refusal(node, key)),
                _ => return Err(element_refusal("set", node)),
            }
        }
    }
    finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        operations,
        limits,
    )
}

/// Remove slides. Several targets are removed from the last to the first,
/// so each one's position still holds when its deletion runs.
pub(crate) fn remove(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Keynote>(file, target, limits.max_expanded_targets())?;
    let mut slides = Vec::with_capacity(nodes.len());
    for node in &nodes {
        if node.item().type_name() != "slide" {
            return Err(element_refusal("remove", node));
        }
        slides.push(node.item().leaf()[0] as usize);
    }
    slides.sort_unstable();
    slides.dedup();
    let mut chain = Chain::new(package.clone());
    let error = |error| leaf_refusal(LeafError::KeynoteDeletion(error), FileFormat::Keynote);
    for (operations, slide) in slides.iter().rev().enumerate() {
        bound(
            LimitKind::Operations,
            operations + 1,
            limits.max_operations(),
        )?;
        let mut edit = chain.package.edit_slide_deletion();
        edit.remove_slide(SlideSelector::index(*slide))
            .map_err(error)?;
        let commit = edit.commit().map_err(error)?;
        let patch = commit.patch().clone();
        chain.push(
            commit.into_package(),
            LeafPatch::KeynoteDeletion(patch),
            true,
        );
    }
    finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        slides.len(),
        limits,
    )
}

/// Move one slide among the presentation's slides.
pub(crate) fn move_to(
    file: &Iwork,
    package: &Package,
    path: &PathArg<'_>,
    placement: &crate::Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let node = node_at::<Keynote>(file, path)?;
    if node.item().type_name() != "slide" {
        return Err(element_refusal("move", &node));
    }
    if matches!(placement, crate::Placement::Into(..)) {
        return Err(refusal(
            Code::UnsupportedCapability,
            "a Keynote slide moves among the presentation's slides",
        ));
    }
    let from = node.item().leaf()[0] as usize;
    let slides = file.tree().item(ROOT).children().len();
    let to = destination(from, slides, placement, |anchor| {
        let anchor = node_at::<Keynote>(file, anchor)?;
        if anchor.item().type_name() == "slide" {
            Ok(anchor.item().leaf()[0] as usize)
        } else {
            Err(refusal(
                Code::InvalidArgument,
                "a slide moves before or after another slide",
            ))
        }
    })?;
    let error = |error| leaf_refusal(LeafError::KeynoteOrder(error), FileFormat::Keynote);
    let mut edit = package.edit_slide_order();
    edit.move_slide(SlideSelector::index(from), Position::new(to))
        .map_err(error)?;
    let commit = edit.commit().map_err(error)?;
    let patch = commit.patch().clone();
    let changed = !patch.is_noop();
    let mut chain = Chain::new(package.clone());
    chain.push(
        commit.into_package(),
        LeafPatch::KeynoteOrder(patch),
        changed,
    );
    finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        1,
        limits,
    )
}
