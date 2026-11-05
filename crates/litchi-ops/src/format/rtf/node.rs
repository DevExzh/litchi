//! Typed generic nodes for RTF documents.
//!
//! A node keeps the handle, its kind, its positional path and a borrowed
//! reference to the native value it stands for. Every value lives in the
//! retained native `Package`, so nodes are `Copy` and build no second model.

use std::borrow::Cow;

use litchi_core::{FileFormat, Position};
use litchi_rtf::field::Field;
use litchi_rtf::list::{List, ListLevel, ListOverride};
use litchi_rtf::style::Style;
use litchi_rtf::table::{Cell, Row, Table};
use litchi_rtf::{Annotation, HeaderFooter, HeaderFooterParagraph, Note, Picture};

use super::Rtf;
use super::children::Children;
use super::spec::Kind;
use crate::error::Result;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, ViewOptions,
};

/// The longest RTF path: `/body/tbl[N]/tr[N]/tc[N]`.
const MAX_STEPS: usize = 4;
const _: () = assert!(MAX_STEPS <= crate::node::INLINE_STEPS);

/// A positional path of at most [`MAX_STEPS`] steps, stored compactly as
/// kinds and positions (record 0287): a node no longer carries four full
/// [`PathStep`]s (224 bytes) and writes them only when its path is read. A
/// style's step is its name, read from the node's style.
#[derive(Clone, Copy)]
pub(crate) struct Steps {
    positions: [usize; MAX_STEPS],
    kinds: [Kind; MAX_STEPS],
    len: u8,
}

impl Steps {
    pub(super) const EMPTY: Self = Self {
        positions: [0; MAX_STEPS],
        kinds: [Kind::Document; MAX_STEPS],
        len: 0,
    };

    /// These steps followed by `kind` at `position`; a path at the longest
    /// RTF depth keeps its steps (no RTF element nests deeper).
    pub(super) const fn then(mut self, kind: Kind, position: Position) -> Self {
        if (self.len as usize) < MAX_STEPS {
            self.kinds[self.len as usize] = kind;
            self.positions[self.len as usize] = position.get();
            self.len += 1;
        }
        self
    }

    const fn len(&self) -> usize {
        self.len as usize
    }
}

/// The native value a non-body node stands for.
#[derive(Clone, Copy)]
pub(crate) enum Item<'f> {
    /// The root, the body or a collection (`/comments`, `/styles`,
    /// `/numbering`).
    Bare,
    /// A borrowed main-story paragraph.
    Paragraph(litchi_rtf::text::Paragraph<'f>),
    /// A borrowed main-story character run.
    Run(litchi_rtf::text::Run<'f>),
    /// A paragraph of a header/footer story.
    StoryParagraph {
        section: usize,
        story: usize,
        index: usize,
        value: &'f HeaderFooterParagraph<'f>,
    },
    /// A top-level table with its native index.
    Table {
        table: usize,
        value: &'f Table<'f>,
    },
    /// A table row with its native table and row indices.
    Row {
        table: usize,
        row: usize,
        value: &'f Row<'f>,
    },
    /// A top-level table cell with its native coordinates and row.
    Cell {
        table: usize,
        row: usize,
        cell: usize,
        value: &'f Cell<'f>,
        parent: &'f Row<'f>,
    },
    /// A header or footer story with its native coordinates.
    Story {
        section: usize,
        story: usize,
        value: &'f HeaderFooter<'f>,
    },
    /// A footnote or endnote: native index and 1-based number of its kind.
    Note {
        index: usize,
        number: usize,
        value: &'f Note<'f>,
    },
    /// A comment with its native index.
    Comment {
        index: usize,
        value: &'f Annotation<'f>,
    },
    /// A main-story `HYPERLINK` field.
    Hyperlink(&'f Field<'f>),
    /// A main-story picture with its native index.
    Picture {
        index: usize,
        value: &'f Picture<'f>,
    },
    Style(&'f Style<'f>),
    AbstractNum(&'f List<'f>),
    Num(&'f ListOverride),
    Level(&'f ListLevel<'f>),
    /// A section with its native index (record 0328).
    Section {
        index: usize,
        value: &'f litchi_rtf::Section<'f>,
    },
    /// A bookmark with its native index.
    Bookmark {
        index: usize,
        value: &'f litchi_rtf::Bookmark<'f>,
    },
    /// A field (or a `TOC` field) with its native index.
    Field {
        index: usize,
        value: &'f Field<'f>,
    },
    /// A tab stop of a main-story paragraph, by value.
    Tab {
        index: usize,
        value: litchi_rtf::TabStop,
    },
    /// A text box: a text-box shape, or a legacy (`\dptxbx`) text box.
    ShapeTextBox {
        value: &'f litchi_rtf::Shape<'f>,
    },
    LegacyTextBox {
        value: &'f litchi_rtf::LegacyTextBox<'f>,
    },
    /// A drawing shape.
    Shape {
        value: &'f litchi_rtf::Shape<'f>,
    },
}

// Record 0287: nodes move through every query; keep this one no larger
// than it is (the unified node is as large as its largest format node).
#[cfg(target_pointer_width = "64")]
const _: () = assert!(size_of::<Node<'static>>() <= 160);

/// A node of an RTF document.
#[derive(Clone, Copy)]
pub struct Node<'f> {
    pub(super) doc: &'f Rtf,
    pub(super) kind: Kind,
    pub(super) item: Item<'f>,
    pub(super) steps: Steps,
}

impl<'f> Node<'f> {
    const fn bare(doc: &'f Rtf, kind: Kind, steps: Steps) -> Self {
        Self {
            doc,
            kind,
            item: Item::Bare,
            steps,
        }
    }

    pub(super) const fn root(doc: &'f Rtf) -> Self {
        Self::bare(doc, Kind::Document, Steps::EMPTY)
    }

    pub(super) const fn body(doc: &'f Rtf) -> Self {
        Self::bare(
            doc,
            Kind::Body,
            Steps::EMPTY.then(Kind::Body, Position::new(0)),
        )
    }

    /// A collection directly under the root (`/comments`, `/styles`,
    /// `/numbering`).
    pub(super) const fn collection(doc: &'f Rtf, kind: Kind) -> Self {
        Self::bare(doc, kind, Steps::EMPTY.then(kind, Position::new(0)))
    }

    pub(super) const fn paragraph(
        doc: &'f Rtf,
        position: Position,
        paragraph: litchi_rtf::text::Paragraph<'f>,
    ) -> Self {
        Self {
            doc,
            kind: Kind::Paragraph,
            item: Item::Paragraph(paragraph),
            steps: Steps::EMPTY
                .then(Kind::Body, Position::new(0))
                .then(Kind::Paragraph, position),
        }
    }

    pub(super) const fn run(
        doc: &'f Rtf,
        paragraph_position: Position,
        run_position: Position,
        run: litchi_rtf::text::Run<'f>,
    ) -> Self {
        Self {
            doc,
            kind: Kind::Run,
            item: Item::Run(run),
            steps: Steps::EMPTY
                .then(Kind::Body, Position::new(0))
                .then(Kind::Paragraph, paragraph_position)
                .then(Kind::Run, run_position),
        }
    }

    /// An inline object (`hyperlink`, `picture`) of main-story paragraph
    /// `paragraph`, without building the paragraph node.
    pub(super) const fn paragraph_child(
        doc: &'f Rtf,
        paragraph: Position,
        kind: Kind,
        item: Item<'f>,
        position: Position,
    ) -> Self {
        Self {
            doc,
            kind,
            item,
            steps: Steps::EMPTY
                .then(Kind::Body, Position::new(0))
                .then(Kind::Paragraph, paragraph)
                .then(kind, position),
        }
    }

    /// A top-level node (`/header[N]`, `/footnote[N]`) for `item`.
    pub(super) fn top(doc: &'f Rtf, kind: Kind, item: Item<'f>, position: Position) -> Self {
        Self {
            item,
            ..Self::bare(doc, kind, Steps::EMPTY.then(kind, position))
        }
    }

    /// A child of `self` with its own kind, value and positional step (a
    /// style's step is its name).
    pub(super) fn child(&self, kind: Kind, item: Item<'f>, position: Position) -> Self {
        Self {
            doc: self.doc,
            kind,
            item,
            steps: self.steps.then(kind, position),
        }
    }

    /// The element kind.
    #[must_use]
    pub const fn element_kind(&self) -> Kind {
        self.kind
    }

    /// The containing main-story paragraph's zero-based position: the
    /// `p[N]` step of a path under `/body/p[N]`.
    #[must_use]
    pub const fn paragraph_position(&self) -> Option<Position> {
        let steps = &self.steps;
        if steps.len() >= 2
            && matches!(steps.kinds[0], Kind::Body)
            && matches!(steps.kinds[1], Kind::Paragraph)
        {
            Some(Position::new(steps.positions[1]))
        } else {
            None
        }
    }

    /// The run's zero-based position within its containing paragraph: the
    /// last step of a main-story run's `/body/p[N]/r[M]` path.
    #[must_use]
    pub const fn run_position(&self) -> Option<Position> {
        let steps = &self.steps;
        if matches!(self.kind, Kind::Run)
            && steps.len() == 3
            && matches!(steps.kinds[2], Kind::Run)
            && self.paragraph_position().is_some()
        {
            Some(Position::new(steps.positions[2]))
        } else {
            None
        }
    }

    /// The borrowed native run retained by this run node.
    #[must_use]
    pub const fn as_run(&self) -> Option<litchi_rtf::text::Run<'f>> {
        self.native_run()
    }

    /// The borrowed native run of a run node.
    pub(super) const fn native_run(&self) -> Option<litchi_rtf::text::Run<'f>> {
        match self.item {
            Item::Run(run) => Some(run),
            _ => None,
        }
    }

    /// The borrowed native paragraph of a main-story paragraph node.
    pub(super) const fn body_paragraph(&self) -> Option<litchi_rtf::text::Paragraph<'f>> {
        match self.item {
            Item::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        }
    }

    /// The node's visible text. Paragraph text remains borrowed when the
    /// native story stores it contiguously; fragmented paragraphs and
    /// multi-paragraph header/footer stories use an owned fallback. Runs,
    /// cells, notes, comments and hyperlink results always borrow.
    pub fn text(&self) -> Result<Option<Cow<'f, str>>> {
        Ok(match self.item {
            Item::Run(run) => Some(Cow::Borrowed(run.text())),
            Item::Paragraph(paragraph) => Some(
                paragraph
                    .as_str()
                    .map_or_else(|| Cow::Owned(paragraph.to_text()), Cow::Borrowed),
            ),
            Item::StoryParagraph { value, .. } => Some(Cow::Borrowed(value.text.as_ref())),
            Item::Cell { value, .. } => Some(Cow::Borrowed(value.text())),
            Item::Story { value, .. } => Some(match value.paragraphs.as_slice() {
                [] => Cow::Borrowed(""),
                [paragraph] => Cow::Borrowed(paragraph.text.as_ref()),
                _ => Cow::Owned(value.text()),
            }),
            Item::Note { value, .. } => Some(Cow::Borrowed(value.content.as_ref())),
            Item::Comment { value, .. } => Some(Cow::Borrowed(value.text.as_ref())),
            Item::Hyperlink(field) => Some(Cow::Borrowed(field.result.as_ref())),
            Item::Field { value, .. } => Some(Cow::Borrowed(value.result.as_ref())),
            Item::Bookmark { value, .. } => Some(Cow::Borrowed(value.content.as_ref())),
            Item::ShapeTextBox { value, .. } | Item::Shape { value, .. } => {
                Some(Cow::Borrowed(value.text.as_ref()))
            },
            Item::LegacyTextBox { value, .. } => Some(Cow::Borrowed(value.text.as_ref())),
            Item::Bare
            | Item::Table { .. }
            | Item::Row { .. }
            | Item::Picture { .. }
            | Item::Style(_)
            | Item::AbstractNum(_)
            | Item::Num(_)
            | Item::Level(_)
            | Item::Section { .. }
            | Item::Tab { .. } => None,
        })
    }

    #[must_use]
    pub fn kind(&self) -> &'static str {
        self.kind.type_name()
    }

    #[must_use]
    pub const fn file_format(&self) -> FileFormat {
        FileFormat::Rtf
    }

    #[must_use]
    pub fn path(&self) -> NodePath<'_> {
        NodePath::inline_with(self.steps.len(), |index| self.step(index)).unwrap_or(NodePath::ROOT)
    }

    /// Path step `index`, written as the path spells it.
    fn step(&self, index: usize) -> PathStep<'f> {
        let position = Position::new(self.steps.positions[index]);
        match self.steps.kinds[index] {
            kind @ (Kind::Document
            | Kind::Body
            | Kind::Comments
            | Kind::Styles
            | Kind::Numbering) => PathStep::new(kind_step(kind)),
            Kind::Style => match self.item {
                Item::Style(style) => PathStep::new(style.name.as_ref()),
                _ => PathStep::at(kind_step(Kind::Style), position),
            },
            kind => PathStep::at(kind_step(kind), position),
        }
    }

    /// The node's own last path step; `None` for the root.
    pub(super) fn last_step(&self) -> Option<PathStep<'f>> {
        self.steps
            .len()
            .checked_sub(1)
            .map(|index| self.step(index))
    }

    #[must_use]
    pub fn positional_path(&self) -> NodePath<'_> {
        self.path()
    }

    pub fn children(&self) -> ViewChildren<'_, 'f> {
        ViewChildren::rtf(Children::of(self))
    }

    /// The 1-based note number a `footnote[N]`/`endnote[N]` selector reads.
    pub(super) const fn number(&self) -> Option<usize> {
        match self.item {
            Item::Note { number, .. } => Some(number),
            _ => None,
        }
    }

    /// The `format` member's entries, in the kind's property-table order.
    pub(super) fn format_props(&self) -> Format<'f> {
        use super::spec;
        let (props, source) = match self.item {
            Item::Paragraph(paragraph) => (
                spec::PARAGRAPH_PROPS,
                Props::Paragraph {
                    doc: self.doc,
                    paragraph,
                },
            ),
            Item::Run(run) => (spec::RUN_PROPS, Props::Run { doc: self.doc, run }),
            Item::StoryParagraph { value, .. } => {
                (spec::PARAGRAPH_PROPS, Props::StoryParagraph(value))
            },
            Item::Table { value, .. } => (spec::TABLE_PROPS, Props::Table(value)),
            Item::Row { value, .. } => (spec::ROW_PROPS, Props::Row(value)),
            Item::Cell {
                value,
                parent,
                cell,
                ..
            } => {
                let previous = match cell.checked_sub(1) {
                    Some(previous) => parent.cells().get(previous).and_then(Cell::right_boundary),
                    None => Some(parent.geometry().left_edge_twips().unwrap_or(0)),
                };
                (spec::CELL_PROPS, Props::Cell { value, previous })
            },
            Item::Story { value, .. } => (spec::STORY_PROPS, Props::Story(value)),
            Item::Note { number, .. } => (spec::NOTE_PROPS, Props::Note(number)),
            Item::Comment { value, .. } => (spec::COMMENT_PROPS, Props::Comment(value)),
            Item::Hyperlink(field) => (spec::HYPERLINK_PROPS, Props::Hyperlink(field)),
            Item::Picture { value, .. } => (spec::PICTURE_PROPS, Props::Picture(value)),
            Item::Style(style) => (
                spec::STYLE_PROPS,
                Props::Style {
                    doc: self.doc,
                    style,
                },
            ),
            Item::AbstractNum(list) => (spec::ABSTRACT_NUM_PROPS, Props::AbstractNum(list)),
            Item::Num(list) => (spec::NUM_PROPS, Props::Num(list)),
            Item::Level(level) => (spec::LEVEL_PROPS, Props::Level(level)),
            Item::Section { value, .. } => (spec::SECTION_PROPS, Props::Section(value)),
            Item::Bookmark { value, .. } => (spec::BOOKMARK_PROPS, Props::Bookmark(value)),
            Item::Field { value, .. } => {
                if self.kind == Kind::Toc {
                    (spec::TOC_PROPS, Props::Toc(value))
                } else {
                    (spec::FIELD_PROPS, Props::Field(value))
                }
            },
            Item::Tab { value, .. } => (spec::TAB_PROPS, Props::Tab(value)),
            Item::ShapeTextBox { value, .. } => (spec::TEXT_BOX_PROPS, Props::ShapeTextBox(value)),
            Item::LegacyTextBox { value, .. } => {
                (spec::TEXT_BOX_PROPS, Props::LegacyTextBox(value))
            },
            Item::Shape { value, .. } => (spec::SHAPE_PROPS, Props::Shape(value)),
            Item::Bare => return Format::empty(),
        };
        Format::new(props, crate::node::Source::Rtf(source)).without(&["text"])
    }
}

/// The path step name of a kind.
pub(super) const fn kind_step(kind: Kind) -> &'static str {
    match kind {
        Kind::Document => "document",
        Kind::Body => "body",
        Kind::Paragraph => "p",
        Kind::Run => "r",
        Kind::Table => "tbl",
        Kind::Row => "tr",
        Kind::Cell => "tc",
        Kind::Header => "header",
        Kind::Footer => "footer",
        Kind::Footnote => "footnote",
        Kind::Endnote => "endnote",
        Kind::Comments => "comments",
        Kind::Comment => "comment",
        Kind::Hyperlink => "hyperlink",
        Kind::Picture => "picture",
        Kind::Styles => "styles",
        Kind::Style => "style",
        Kind::Numbering => "numbering",
        Kind::AbstractNum => "abstractNum",
        Kind::Num => "num",
        Kind::Level => "level",
        Kind::Section => "section",
        Kind::Bookmark => "bookmark",
        Kind::Field => "field",
        Kind::Toc => "toc",
        Kind::Tab => "tab",
        Kind::TextBox => "textbox",
        Kind::Shape => "shape",
    }
}

/// The values of an RTF node's `format` member, in its kind's property-table
/// order. Every value borrows the retained native package.
#[derive(Debug, Clone, Copy)]
pub(crate) enum Props<'a> {
    Paragraph {
        doc: &'a Rtf,
        paragraph: litchi_rtf::text::Paragraph<'a>,
    },
    Run {
        doc: &'a Rtf,
        run: litchi_rtf::text::Run<'a>,
    },
    StoryParagraph(&'a HeaderFooterParagraph<'a>),
    Table(&'a Table<'a>),
    Row(&'a Row<'a>),
    Cell {
        value: &'a Cell<'a>,
        /// The right boundary of the preceding cell, or the row's left edge.
        previous: Option<i32>,
    },
    Story(&'a HeaderFooter<'a>),
    Note(usize),
    Comment(&'a Annotation<'a>),
    Hyperlink(&'a Field<'a>),
    Picture(&'a Picture<'a>),
    Style {
        doc: &'a Rtf,
        style: &'a Style<'a>,
    },
    AbstractNum(&'a List<'a>),
    Num(&'a ListOverride),
    Level(&'a ListLevel<'a>),
    Section(&'a litchi_rtf::Section<'a>),
    Bookmark(&'a litchi_rtf::Bookmark<'a>),
    Field(&'a Field<'a>),
    Toc(&'a Field<'a>),
    Tab(litchi_rtf::TabStop),
    ShapeTextBox(&'a litchi_rtf::Shape<'a>),
    LegacyTextBox(&'a litchi_rtf::LegacyTextBox<'a>),
    Shape(&'a litchi_rtf::Shape<'a>),
}

impl<'a> Props<'a> {
    /// The value at `index` of the kind's property table, if the node has it.
    #[allow(
        clippy::too_many_lines,
        reason = "one flat projection per native kind keeps the table order visible"
    )]
    pub(crate) fn get(&self, index: usize) -> Option<crate::node::Value<'a>> {
        use crate::node::Value;
        use crate::value::{Color, Length};
        let int = |value: usize| i64::try_from(value).ok().map(Value::Int);
        match *self {
            Self::Paragraph { doc, paragraph } => {
                let format = paragraph.format();
                let layout = format.layout();
                match index {
                    0 => Some(Value::Text(
                        paragraph
                            .as_str()
                            .map_or_else(|| Cow::Owned(paragraph.to_text()), Cow::Borrowed),
                    )),
                    1 => Some(Value::Token(super::alignment_name(format.alignment()))),
                    2 => Some(Value::Length(Length::from_twips(layout.space_before()))),
                    3 => Some(Value::Length(Length::from_twips(layout.space_after()))),
                    4 => Some(Value::Length(Length::from_twips(layout.left_indent()))),
                    5 => Some(Value::Length(Length::from_twips(layout.right_indent()))),
                    6 => Some(Value::Length(Length::from_twips(
                        layout.first_line_indent(),
                    ))),
                    7 => Some(Value::Bool(layout.keep_with_next())),
                    8 => Some(Value::Bool(layout.keep_together())),
                    9 => Some(Value::Bool(layout.page_break_before())),
                    10 => format.style().and_then(|id| {
                        super::stories::style_name(doc, litchi_rtf::style::Kind::Paragraph, id)
                            .map(|name| Value::Text(Cow::Borrowed(name)))
                    }),
                    11 => format.list_override().map(|id| Value::Int(i64::from(id))),
                    12 => format
                        .list_level()
                        .map(|level| Value::Int(i64::from(level))),
                    _ => None,
                }
            },
            Self::Run { doc, run } => {
                let format = run.format();
                match index {
                    0 => Some(Value::Text(Cow::Borrowed(run.text()))),
                    1 => format
                        .font()
                        .map(|font| Value::Text(Cow::Borrowed(font.name()))),
                    2 => Some(Value::Length(Length::from_half_points(i32::from(
                        format.size().get(),
                    )))),
                    3 => Some(Value::Bool(format.bold())),
                    4 => Some(Value::Bool(format.italic())),
                    5 => format.foreground().and_then(|value| match value {
                        litchi_rtf::color::Value::Automatic => Some(Value::Color(Color::AUTO)),
                        litchi_rtf::color::Value::Rgb(color) => Some(Value::Color(Color::rgb(
                            color.red(),
                            color.green(),
                            color.blue(),
                        ))),
                        _ => None,
                    }),
                    6 => Some(Value::Token(super::underline_name(format.underline()))),
                    7 => Some(Value::Bool(format.strike())),
                    8 => Some(Value::Bool(format.double_strike())),
                    9 => Some(Value::Bool(format.hidden())),
                    10 => Some(Value::Bool(format.all_caps())),
                    11 => Some(Value::Bool(format.small_caps())),
                    12 => Some(Value::Bool(format.outline())),
                    13 => format.character_style().and_then(|id| {
                        super::stories::style_name(doc, litchi_rtf::style::Kind::Character, id)
                            .map(|name| Value::Text(Cow::Borrowed(name)))
                    }),
                    _ => None,
                }
            },
            Self::StoryParagraph(paragraph) => match index {
                0 => Some(Value::Text(Cow::Borrowed(paragraph.text.as_ref()))),
                _ => None,
            },
            Self::Table(table) => match index {
                0 => int(table.row_count()),
                1 => int(table.rows().iter().map(Row::cell_count).max().unwrap_or(0)),
                _ => None,
            },
            Self::Row(row) => match index {
                0 => Some(Value::Bool(row.layout().header)),
                1 => Some(Value::Bool(row.layout().keep_together)),
                2 => match row.geometry().height() {
                    litchi_rtf::table::TableRowHeight::Automatic => None,
                    litchi_rtf::table::TableRowHeight::Minimum(twips)
                    | litchi_rtf::table::TableRowHeight::Exact(twips) => {
                        Some(Value::Length(Length::from_twips(i32::from(twips))))
                    },
                },
                _ => None,
            },
            Self::Cell {
                value: cell,
                previous,
            } => match index {
                0 => Some(Value::Text(Cow::Borrowed(cell.text()))),
                1 => cell
                    .preferred_width()
                    .and_then(|width| match width.unit() {
                        litchi_rtf::table::TablePreferredWidthUnit::Twips => {
                            width.value().map(i32::from)
                        },
                        _ => None,
                    })
                    .or_else(|| {
                        cell.right_boundary()?
                            .checked_sub(previous?)
                            .filter(|width| *width > 0)
                    })
                    .map(|twips| Value::Length(Length::from_twips(twips))),
                2 => cell.layout().vertical_alignment.map(|alignment| {
                    Value::Token(match alignment {
                        litchi_rtf::table::TableCellVerticalAlignment::Top => "top",
                        litchi_rtf::table::TableCellVerticalAlignment::Center => "center",
                        litchi_rtf::table::TableCellVerticalAlignment::Bottom => "bottom",
                    })
                }),
                3 => cell.merge().vertical.map(merge_name),
                4 => cell.merge().horizontal.map(merge_name),
                _ => None,
            },
            Self::Story(story) => match index {
                1 => Some(Value::Token(super::stories::story_type(story.header_type))),
                _ => None,
            },
            Self::Note(number) => match index {
                1 => int(number),
                _ => None,
            },
            Self::Comment(comment) => match index {
                1 => Some(Value::Text(Cow::Borrowed(comment.author.as_ref()))),
                2 => Some(Value::Text(Cow::Borrowed(comment.initials.as_ref()))),
                3 => comment
                    .date
                    .as_deref()
                    .map(|date| Value::Text(Cow::Borrowed(date))),
                _ => None,
            },
            Self::Hyperlink(field) => {
                if index == 0 {
                    return Some(Value::Text(Cow::Borrowed(field.result.as_ref())));
                }
                let link = field.hyperlink()?;
                let text = |value: Option<&str>| {
                    value.map(|value| Value::Text(Cow::Owned(value.to_owned())))
                };
                match index {
                    1 => text(link.external_target()),
                    2 => text(link.bookmark()),
                    3 => text(link.screen_tip()),
                    4 => text(link.target_frame()),
                    _ => None,
                }
            },
            Self::Picture(picture) => match index {
                0 => picture_extent(picture.goal_width, picture.scale_x)
                    .or_else(|| picture_extent(picture.width, None))
                    .map(Value::Length),
                1 => picture_extent(picture.goal_height, picture.scale_y)
                    .or_else(|| picture_extent(picture.height, None))
                    .map(Value::Length),
                2 => Some(Value::Token(match picture.image_type {
                    litchi_rtf::ImageType::Png => "png",
                    litchi_rtf::ImageType::Jpeg => "jpeg",
                    litchi_rtf::ImageType::Emf => "emf",
                    litchi_rtf::ImageType::Wmf => "wmf",
                    litchi_rtf::ImageType::Pict => "pict",
                    litchi_rtf::ImageType::Dib => "dib",
                    litchi_rtf::ImageType::Unknown => "unknown",
                })),
                _ => None,
            },
            Self::Style { doc, style } => {
                let named = |id: Option<u16>| {
                    id.and_then(|id| super::stories::style_name(doc, style.style_type, id))
                        .map(|name| Value::Text(Cow::Borrowed(name)))
                };
                match index {
                    0 => Some(Value::Text(Cow::Borrowed(style.name.as_ref()))),
                    1 => Some(Value::Token(match style.style_type {
                        litchi_rtf::style::Kind::Paragraph => "paragraph",
                        litchi_rtf::style::Kind::Character => "character",
                        litchi_rtf::style::Kind::Section => "section",
                        litchi_rtf::style::Kind::Table => "table",
                    })),
                    2 => Some(Value::Int(i64::from(style.id))),
                    3 => named(style.based_on),
                    4 => named(style.next_style),
                    _ => None,
                }
            },
            Self::AbstractNum(list) => match index {
                0 => Some(Value::Int(i64::from(list.id))),
                1 => {
                    (!list.name.is_empty()).then(|| Value::Text(Cow::Borrowed(list.name.as_ref())))
                },
                2 => int(list.levels.len()),
                _ => None,
            },
            Self::Num(list) => match index {
                0 => Some(Value::Int(i64::from(list.index))),
                1 => Some(Value::Int(i64::from(list.list_id))),
                _ => None,
            },
            Self::Level(level) => match index {
                0 => Some(Value::Int(i64::from(level.level))),
                1 => Some(Value::Token(number_format(level.level_type))),
                2 => Some(Value::Text(Cow::Borrowed(level.number_text.as_ref()))),
                3 => Some(Value::Int(i64::from(level.start_at))),
                _ => None,
            },
            Self::Section(section) => super::parts::section_value(&section.properties, index),
            Self::Bookmark(bookmark) => match index {
                0 => Some(Value::Text(Cow::Borrowed(bookmark.name.as_ref()))),
                _ => None,
            },
            Self::Field(field) => super::parts::field_value(field, index),
            Self::Toc(field) => super::parts::toc_value(field, index),
            Self::Tab(stop) => super::parts::tab_value(stop, index),
            Self::ShapeTextBox(shape) => match index {
                0 => twips_length(shape.geometry.width),
                1 => twips_length(shape.geometry.height),
                _ => None,
            },
            Self::LegacyTextBox(text_box) => match index {
                0 => text_box.width.and_then(twips_length),
                1 => text_box.height.and_then(twips_length),
                _ => None,
            },
            Self::Shape(shape) => super::parts::shape_value(shape, index),
        }
    }
}

/// A nonnegative twips extent as a length; a negative one reads as absent.
fn twips_length(twips: i32) -> Option<crate::node::Value<'static>> {
    (twips >= 0).then(|| crate::node::Value::Length(crate::value::Length::from_twips(twips)))
}

const fn merge_name(role: litchi_rtf::table::TableCellMergeRole) -> crate::node::Value<'static> {
    crate::node::Value::Token(match role {
        litchi_rtf::table::TableCellMergeRole::First => "restart",
        litchi_rtf::table::TableCellMergeRole::Continuation => "continue",
    })
}

/// A picture extent from twips and an optional percentage scale.
fn picture_extent(twips: Option<i32>, scale: Option<i32>) -> Option<crate::value::Length> {
    let twips = i64::from(twips?);
    let scaled = match scale {
        Some(percent) => twips.checked_mul(i64::from(percent))?.checked_div(100)?,
        None => twips,
    };
    i32::try_from(scaled)
        .ok()
        .map(crate::value::Length::from_twips)
}

const fn number_format(kind: litchi_rtf::list::ListLevelType) -> &'static str {
    use litchi_rtf::list::ListLevelType as T;
    match kind {
        T::Decimal => "decimal",
        T::UpperRoman => "upperRoman",
        T::LowerRoman => "lowerRoman",
        T::UpperLetter => "upperLetter",
        T::LowerLetter => "lowerLetter",
        T::Ordinal => "ordinal",
        T::CardinalText => "cardinalText",
        T::OrdinalText => "ordinalText",
        T::Bullet => "bullet",
        T::None => "none",
        T::Other(_) => "other",
    }
}

impl std::fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("rtf::Node")
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        self.kind()
    }

    fn file_format(&self) -> FileFormat {
        self.file_format()
    }

    fn path(&self) -> NodePath<'_> {
        self.path()
    }

    fn positional_path(&self) -> NodePath<'_> {
        self.positional_path()
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        Node::text(self)
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        Ok(Node::text(self)?.map(Preview::Prefix))
    }

    fn style(&self) -> Result<Option<Cow<'_, str>>> {
        Ok(self
            .body_paragraph()
            .and_then(|paragraph| paragraph.format().style())
            .and_then(|id| {
                super::stories::style_name(self.doc, litchi_rtf::style::Kind::Paragraph, id)
            })
            .map(Cow::Borrowed))
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        let package = &self.doc.0.package;
        let count = match (self.kind, self.body_paragraph()) {
            (Kind::Body, _) => package
                .paragraph_count()
                .saturating_add(package.tables().len()),
            (Kind::Paragraph, Some(paragraph)) => {
                let anchors = super::stories::anchor_range(paragraph);
                paragraph
                    .runs()
                    .count()
                    .saturating_add(super::stories::hyperlinks(self.doc, anchors.clone()).count())
                    .saturating_add(super::stories::pictures(self.doc, anchors).count())
            },
            _ => Children::of(self).count(),
        };
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    fn format(&self) -> Result<Format<'_>> {
        Ok(self.format_props())
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        self.children()
    }
}
