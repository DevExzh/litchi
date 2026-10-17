//! Typed generic nodes for legacy Word documents.
//!
//! A node keeps the handle, its kind, its positional path and the native
//! value it stands for. Main-story paragraphs stay lazy: a point lookup or a
//! body walk carries the borrowed paragraph text, and the native paragraph
//! (properties and runs) is built only when the node's format or runs are
//! read. Every other value is the one native element the path selected,
//! shared behind an `Arc` so cloning a node never copies document content.

use std::borrow::Cow;
use std::sync::{Arc, OnceLock};

use litchi_core::{FileFormat, Position};

use super::Doc;
use super::children::Children;
use super::spec::{self, Kind};
use crate::error::Result;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};
use crate::value::{Color, Length};

/// The longest DOC path: `/body/tbl[N]/tr[N]/tc[N]/p[N]/r[N]`.
const MAX_STEPS: usize = 6;

/// A positional path of at most [`MAX_STEPS`] steps, stored compactly as
/// kinds and positions (record 0287): a node no longer carries six full
/// [`PathStep`]s (336 bytes) and writes them only when its path is read.
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

    /// These steps followed by `kind` at `position`; a path already at the
    /// longest DOC depth keeps its steps (no DOC element nests deeper).
    pub(super) const fn then(mut self, kind: Kind, position: Position) -> Self {
        if (self.len as usize) < MAX_STEPS {
            self.kinds[self.len as usize] = kind;
            self.positions[self.len as usize] = position.get();
            self.len += 1;
        }
        self
    }

    pub(super) const fn len(&self) -> usize {
        self.len as usize
    }

    /// Step `index`, written as the path spells it.
    fn step<'f>(&self, index: usize, doc: &'f Doc) -> PathStep<'f> {
        step_of(self.kinds[index], Position::new(self.positions[index]), doc)
    }

    /// The last step, `None` at the root.
    pub(super) fn last<'f>(&self, doc: &'f Doc) -> Option<PathStep<'f>> {
        self.len().checked_sub(1).map(|index| self.step(index, doc))
    }
}

/// The path step of `kind` at `position`: the collections are bare, and a
/// style is its identifier (`/styles/Heading1`; record 0324).
fn step_of(kind: Kind, position: Position, doc: &Doc) -> PathStep<'_> {
    match kind {
        Kind::Document => PathStep::new("document"),
        Kind::Body => PathStep::new("body"),
        Kind::Comments => PathStep::new("comments"),
        Kind::Styles => PathStep::new("styles"),
        Kind::Numbering => PathStep::new("numbering"),
        Kind::Style => match u16::try_from(position.get())
            .ok()
            .and_then(|istd| doc.style_id(istd))
        {
            Some(id) => PathStep::new(id),
            None => PathStep::at(kind_step(kind), position),
        },
        kind => PathStep::at(kind_step(kind), position),
    }
}

/// A header, footer, note or comment story.
pub(crate) enum Story {
    HeaderFooter(litchi_doc::HeaderFooter),
    Note(litchi_doc::Footnote),
    Comment(litchi_doc::Comment),
}

impl Story {
    pub(super) fn paragraphs(&self) -> &[litchi_doc::Paragraph] {
        match self {
            Self::HeaderFooter(story) => &story.paragraphs,
            Self::Note(story) => &story.paragraphs,
            Self::Comment(story) => &story.paragraphs,
        }
    }
}

/// The native value a node stands for.
#[derive(Clone)]
pub(crate) enum Item<'f> {
    /// The root, the body or the comment collection.
    Bare,
    /// An ordinary main-story paragraph (`/body/p[N]`).
    BodyParagraph {
        position: Position,
        /// The paragraph text, when the node came from the body walk.
        text: Option<&'f str>,
        /// The native paragraph, built on the first format or run read.
        built: OnceLock<Arc<litchi_doc::Paragraph>>,
    },
    /// A paragraph of a table cell or a story.
    Paragraph(Arc<litchi_doc::Paragraph>),
    /// A run. `span` is its UTF-16 range in a main-story paragraph's text,
    /// the leaf edit target; runs of other paragraphs have none.
    Run {
        run: Arc<litchi_doc::Run>,
        body: Option<(Position, u32, u32)>,
    },
    Table(Arc<litchi_doc::Table>),
    Row(Arc<litchi_doc::Row>),
    Cell(Arc<litchi_doc::Cell>),
    Story {
        story: Arc<Story>,
        /// The note's 1-based ordinal (`footnote[N]`); 0 for other stories.
        number: usize,
    },
    /// A style (`istd`) and its direct properties, parsed on first read.
    Style(u16, OnceLock<Arc<super::parts::StyleView>>),
    /// A list definition, by index into `PlfLst`.
    AbstractNum(usize),
    /// A list instance, by index into `PlfLfo`.
    Num(usize),
    /// Level `level` of definition `definition`, with its properties.
    Level {
        definition: usize,
        level: &'f litchi_doc::ListLevel,
        view: OnceLock<Arc<super::parts::LevelView>>,
    },
    Hyperlink(Arc<super::marks::Link>),
    Bookmark(Arc<litchi_doc::Bookmark>),
    Field(Arc<litchi_doc::FieldText>),
    Toc(Arc<litchi_doc::FieldText>),
    Section(&'f litchi_doc::section::Section),
    /// A tab stop of a paragraph.
    Tab(litchi_doc::parts::pap::TabStop),
    TextBox(Arc<super::marks::TextBox>),
    Shape(&'f litchi_doc::parts::spa::ShapeAnchor),
}

// Record 0287: nodes move through every query; keep this one no larger
// than it is (the unified node is as large as its largest format node).
#[cfg(target_pointer_width = "64")]
const _: () = assert!(size_of::<Node<'static>>() <= 144);

/// A node of a DOC document.
#[derive(Clone)]
pub struct Node<'f> {
    pub(super) doc: &'f Doc,
    pub(super) kind: Kind,
    pub(super) item: Item<'f>,
    pub(super) steps: Steps,
    /// The spelled-out steps of a path deeper than [`NodePath`] carries by
    /// value, built on the first read of such a path.
    paths: OnceLock<Box<[PathStep<'f>]>>,
}

impl<'f> Node<'f> {
    pub(super) const fn root(doc: &'f Doc) -> Self {
        Self {
            doc,
            kind: Kind::Document,
            item: Item::Bare,
            steps: Steps::EMPTY,
            paths: OnceLock::new(),
        }
    }

    pub(super) const fn body(doc: &'f Doc) -> Self {
        Self {
            doc,
            kind: Kind::Body,
            item: Item::Bare,
            steps: Steps::EMPTY.then(Kind::Body, Position::new(0)),
            paths: OnceLock::new(),
        }
    }

    pub(super) const fn comments(doc: &'f Doc) -> Self {
        Self {
            doc,
            kind: Kind::Comments,
            item: Item::Bare,
            steps: Steps::EMPTY.then(Kind::Comments, Position::new(0)),
            paths: OnceLock::new(),
        }
    }

    pub(super) fn paragraph(doc: &'f Doc, position: Position) -> Self {
        Self::body_paragraph(doc, position, None)
    }

    pub(super) fn paragraph_with_text(doc: &'f Doc, position: Position, text: &'f str) -> Self {
        Self::body_paragraph(doc, position, Some(text))
    }

    fn body_paragraph(doc: &'f Doc, position: Position, text: Option<&'f str>) -> Self {
        Self {
            doc,
            kind: Kind::Paragraph,
            item: Item::BodyParagraph {
                position,
                text,
                built: OnceLock::new(),
            },
            steps: Steps::EMPTY
                .then(Kind::Body, Position::new(0))
                .then(Kind::Paragraph, position),
            paths: OnceLock::new(),
        }
    }

    /// A child of `self` with its own kind, value and last step.
    pub(super) fn child(&self, kind: Kind, item: Item<'f>, position: Position) -> Self {
        Self {
            doc: self.doc,
            kind,
            item,
            steps: self.steps.then(kind, position),
            paths: OnceLock::new(),
        }
    }

    /// A top-level story node (`/header[N]`, `/footnote[N]`).
    pub(super) fn story(doc: &'f Doc, kind: Kind, story: Arc<Story>, position: Position) -> Self {
        let number = match kind {
            Kind::Footnote | Kind::Endnote => position.get().saturating_add(1),
            _ => 0,
        };
        Self {
            doc,
            kind,
            item: Item::Story { story, number },
            steps: Steps::EMPTY.then(kind, position),
            paths: OnceLock::new(),
        }
    }

    /// The element kind.
    #[must_use]
    pub const fn element_kind(&self) -> Kind {
        self.kind
    }

    /// The zero-based position of an ordinary main-story paragraph, the
    /// identity shared by reads and the leaf body editor. `None` for every
    /// other node, paragraphs of cells and stories included.
    #[must_use]
    pub const fn paragraph_position(&self) -> Option<Position> {
        match self.item {
            Item::BodyParagraph { position, .. } => Some(position),
            _ => None,
        }
    }

    /// For a run of an ordinary main-story paragraph: the paragraph's
    /// position and the run's UTF-16 range in the paragraph's text.
    #[must_use]
    pub(crate) const fn body_run(&self) -> Option<(Position, u32, u32)> {
        match self.item {
            Item::Run { body, .. } => body,
            _ => None,
        }
    }

    /// The native run of a run node.
    pub(crate) fn run(&self) -> Option<&litchi_doc::Run> {
        match &self.item {
            Item::Run { run, .. } => Some(run),
            _ => None,
        }
    }

    /// The paragraph text for a paragraph node.
    ///
    /// # Errors
    ///
    /// Returns a typed DOC error when the paragraph's story cannot be read.
    pub fn paragraph_text(&self) -> Result<Option<Cow<'_, str>>> {
        self.doc.check_source()?;
        match &self.item {
            Item::BodyParagraph { position, text, .. } => {
                if let Some(text) = text {
                    return Ok(Some(Cow::Borrowed(text)));
                }
                Ok(Some(Cow::Borrowed(
                    self.doc.body_paragraph_at(*position)?.ok_or_else(|| {
                        crate::Error::new(
                            crate::Code::PathNotFound,
                            "the DOC paragraph is outside the main story",
                        )
                        .with_format(FileFormat::Doc)
                    })?,
                )))
            },
            Item::Paragraph(paragraph) => Ok(Some(visible(paragraph.text()?))),
            _ => Ok(None),
        }
    }

    /// The native paragraph of a paragraph node, built on first use for a
    /// main-story paragraph.
    pub(super) fn native_paragraph(&self) -> Result<Option<&litchi_doc::Paragraph>> {
        match &self.item {
            Item::BodyParagraph {
                position, built, ..
            } => {
                if let Some(paragraph) = built.get() {
                    return Ok(Some(paragraph));
                }
                let paragraph = self
                    .doc
                    .reader()?
                    .body_paragraph(*position)?
                    .ok_or_else(|| {
                        crate::Error::new(
                            crate::Code::PathNotFound,
                            "the DOC paragraph is outside the main story",
                        )
                        .with_format(FileFormat::Doc)
                    })?;
                let _ = built.set(Arc::new(paragraph));
                Ok(built.get().map(AsRef::as_ref))
            },
            Item::Paragraph(paragraph) => Ok(Some(paragraph)),
            _ => Ok(None),
        }
    }

    /// The node's text: a paragraph, run, cell or story; `None` otherwise.
    ///
    /// # Errors
    ///
    /// Returns a typed DOC error when the text's story cannot be read.
    pub fn text(&self) -> Result<Option<Cow<'_, str>>> {
        match &self.item {
            Item::BodyParagraph { .. } | Item::Paragraph(_) => self.paragraph_text(),
            Item::Run { run, .. } => Ok(Some(visible(run.text()?))),
            Item::Cell(cell) => Ok(Some(visible(cell.text()?))),
            Item::Story { story, .. } => {
                let mut text = String::new();
                for (index, paragraph) in story.paragraphs().iter().enumerate() {
                    if index > 0 {
                        text.push('\n');
                    }
                    text.push_str(&visible(paragraph.text()?));
                }
                // A header or footer story ends with a guard paragraph mark.
                let end = text.trim_end_matches('\n').len();
                text.truncate(end);
                Ok(Some(Cow::Owned(text)))
            },
            Item::Hyperlink(link) => Ok(Some(link.text())),
            Item::Field(field) | Item::Toc(field) => {
                Ok(field.result.as_deref().map(super::marks::result_text))
            },
            Item::TextBox(text_box) => Ok(Some(visible(&text_box.text_box.text))),
            _ => Ok(None),
        }
    }

    #[must_use]
    pub fn kind(&self) -> &'static str {
        self.kind.type_name()
    }

    #[must_use]
    pub const fn file_format(&self) -> FileFormat {
        FileFormat::Doc
    }

    #[must_use]
    pub fn path(&self) -> NodePath<'_> {
        let len = self.steps.len();
        let doc = self.doc;
        if let Some(path) = NodePath::inline_with(len, |index| self.steps.step(index, doc)) {
            return path;
        }
        NodePath::new(
            self.paths
                .get_or_init(|| (0..len).map(|index| self.steps.step(index, doc)).collect()),
        )
    }

    /// The node's own last path step; `None` for the root.
    pub(super) fn last_step(&self) -> Option<PathStep<'f>> {
        self.steps.last(self.doc)
    }

    /// The body position of the ordinary main-story paragraph this node is,
    /// or is a run, hyperlink or tab stop of, read from the node's path.
    pub(crate) fn body_paragraph_of(&self) -> Option<Position> {
        let steps = &self.steps;
        (steps.len() >= 2 && steps.kinds[0] == Kind::Body && steps.kinds[1] == Kind::Paragraph)
            .then(|| Position::new(steps.positions[1]))
    }

    /// The node's own position among its kind's siblings.
    pub(crate) fn own_position(&self) -> Option<Position> {
        let len = self.steps.len();
        len.checked_sub(1)
            .map(|last| Position::new(self.steps.positions[last]))
    }

    /// The node's native value.
    pub(crate) const fn item(&self) -> &Item<'f> {
        &self.item
    }

    /// A top-level part or mark node (`/styles`, `/bookmark[N]`).
    pub(super) fn top(doc: &'f Doc, kind: Kind, item: Item<'f>, position: Position) -> Self {
        Self {
            doc,
            kind,
            item,
            steps: Steps::EMPTY.then(kind, position),
            paths: OnceLock::new(),
        }
    }

    #[must_use]
    pub fn positional_path(&self) -> NodePath<'_> {
        self.path()
    }

    pub fn children(&self) -> ViewChildren<'_, 'f> {
        ViewChildren::doc(Children::of(self))
    }

    /// The note number a `footnote[N]`/`endnote[N]` selector compares.
    pub(super) const fn number(&self) -> Option<usize> {
        match self.item {
            Item::Story { number, .. } if number > 0 => Some(number),
            _ => None,
        }
    }

    pub(super) fn format_props(&self) -> Result<Format<'_>> {
        let doc = self.doc;
        let (props, source) = match &self.item {
            Item::BodyParagraph { .. } | Item::Paragraph(_) => match self.native_paragraph()? {
                Some(paragraph) => (
                    spec::PARAGRAPH_PROPS,
                    NodeProps::Paragraph { paragraph, doc },
                ),
                None => return Ok(Format::empty()),
            },
            Item::Run { run, .. } => (spec::RUN_PROPS, NodeProps::Run { run, doc: self.doc }),
            Item::Table(table) => (spec::TABLE_PROPS, NodeProps::Table(table)),
            Item::Row(row) => (spec::ROW_PROPS, NodeProps::Row(row)),
            Item::Cell(cell) => (spec::CELL_PROPS, NodeProps::Cell(cell)),
            Item::Story { story, number } => match story.as_ref() {
                Story::HeaderFooter(story) => (spec::STORY_PROPS, NodeProps::HeaderFooter(story)),
                Story::Note(_) => (spec::NOTE_PROPS, NodeProps::Note(*number)),
                Story::Comment(story) => (spec::COMMENT_PROPS, NodeProps::Comment(story)),
            },
            Item::Bare => match self.kind {
                Kind::Styles => (
                    spec::STYLES_PROPS,
                    NodeProps::Count(
                        doc.reader()?
                            .stylesheet()
                            .map_or(0, |sheet| sheet.styles().iter().flatten().count()),
                        0,
                    ),
                ),
                Kind::Numbering => (
                    spec::NUMBERING_PROPS,
                    NodeProps::Count(
                        super::parts::structures(doc).len(),
                        super::parts::num_count(doc),
                    ),
                ),
                _ => return Ok(Format::empty()),
            },
            Item::Style(istd, view) => {
                let view = match view.get() {
                    Some(view) => view,
                    None => {
                        let Some(style) = doc
                            .reader()?
                            .stylesheet()
                            .and_then(|sheet| sheet.get(*istd))
                        else {
                            return Ok(Format::empty());
                        };
                        let _ = view.set(Arc::new(super::parts::StyleView::of(style)));
                        view.get().ok_or_else(|| {
                            crate::Error::new(crate::Code::InternalError, "DOC style view missing")
                        })?
                    },
                };
                (
                    spec::STYLE_PROPS,
                    NodeProps::Style {
                        doc,
                        istd: *istd,
                        view,
                    },
                )
            },
            Item::AbstractNum(at) => (
                spec::ABSTRACT_NUM_PROPS,
                NodeProps::AbstractNum { doc, at: *at },
            ),
            Item::Num(at) => (spec::NUM_PROPS, NodeProps::Num { doc, at: *at }),
            Item::Level {
                definition,
                level,
                view,
            } => {
                let view = view.get_or_init(|| {
                    Arc::new(super::parts::LevelView::of(doc, *definition, level.level))
                });
                (
                    spec::LEVEL_PROPS,
                    NodeProps::Level {
                        doc,
                        definition: *definition,
                        level,
                        view,
                    },
                )
            },
            Item::Hyperlink(link) => (spec::HYPERLINK_PROPS, NodeProps::Hyperlink { doc, link }),
            Item::Bookmark(bookmark) => (spec::BOOKMARK_PROPS, NodeProps::Bookmark(bookmark)),
            Item::Field(field) => (spec::FIELD_PROPS, NodeProps::Field(field)),
            Item::Toc(field) => (spec::TOC_PROPS, NodeProps::Toc(field)),
            Item::Section(section) => (spec::SECTION_PROPS, NodeProps::Section(section)),
            Item::Tab(stop) => (spec::TAB_PROPS, NodeProps::Tab(stop)),
            Item::TextBox(text_box) => (spec::TEXT_BOX_PROPS, NodeProps::TextBox(text_box)),
            Item::Shape(anchor) => (spec::SHAPE_PROPS, NodeProps::Shape(anchor)),
        };
        Ok(Format::new(props, Source::Doc(source)).without(&["text"]))
    }
}

/// The path step name of a kind.
const fn kind_step(kind: Kind) -> &'static str {
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
        Kind::Styles => "styles",
        Kind::Style => "style",
        Kind::Numbering => "numbering",
        Kind::AbstractNum => "abstractNum",
        Kind::Num => "num",
        Kind::Level => "level",
        Kind::Hyperlink => "hyperlink",
        Kind::Bookmark => "bookmark",
        Kind::Field => "field",
        Kind::Toc => "toc",
        Kind::Section => "section",
        Kind::Tab => "tab",
        Kind::TextBox => "textbox",
        Kind::Shape => "shape",
    }
}

/// Text without DOC control characters (marks, field delimiters, note and
/// annotation references), tabs kept.
fn visible(text: &str) -> Cow<'_, str> {
    if text
        .chars()
        .any(|character| character.is_control() && character != '\t')
    {
        Cow::Owned(
            text.chars()
                .filter(|character| !character.is_control() || *character == '\t')
                .collect(),
        )
    } else {
        Cow::Borrowed(text)
    }
}

/// The values of a DOC node's `format` member, in its kind's property-table
/// order.
#[derive(Debug, Clone, Copy)]
pub(crate) enum NodeProps<'a> {
    Paragraph {
        paragraph: &'a litchi_doc::Paragraph,
        doc: &'a Doc,
    },
    Run {
        run: &'a litchi_doc::Run,
        doc: &'a Doc,
    },
    Table(&'a litchi_doc::Table),
    Row(&'a litchi_doc::Row),
    Cell(&'a litchi_doc::Cell),
    HeaderFooter(&'a litchi_doc::HeaderFooter),
    Note(usize),
    Comment(&'a litchi_doc::Comment),
    /// `/styles` (style count) or `/numbering` (definition and instance
    /// counts).
    Count(usize, usize),
    Style {
        doc: &'a Doc,
        istd: u16,
        view: &'a super::parts::StyleView,
    },
    AbstractNum {
        doc: &'a Doc,
        at: usize,
    },
    Num {
        doc: &'a Doc,
        at: usize,
    },
    Level {
        doc: &'a Doc,
        definition: usize,
        level: &'a litchi_doc::ListLevel,
        view: &'a super::parts::LevelView,
    },
    Hyperlink {
        doc: &'a Doc,
        link: &'a super::marks::Link,
    },
    Bookmark(&'a litchi_doc::Bookmark),
    Field(&'a litchi_doc::FieldText),
    Toc(&'a litchi_doc::FieldText),
    Section(&'a litchi_doc::section::Section),
    Tab(&'a litchi_doc::parts::pap::TabStop),
    TextBox(&'a super::marks::TextBox),
    Shape(&'a litchi_doc::parts::spa::ShapeAnchor),
}

impl<'a> NodeProps<'a> {
    /// The value at `index` of the kind's property table, if the node has
    /// it. Text entries are node members and are read through `text`.
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        match *self {
            Self::Paragraph { paragraph, doc } => {
                use super::parts::{ParagraphKey as K, paragraph_value};
                let properties = paragraph.properties();
                match index {
                    1 => Some(Value::Token(spec::alignment_name(properties.justification))),
                    2 => properties
                        .space_before
                        .map(|twips| Value::Length(Length::from_twips(i32::from(twips)))),
                    3 => properties
                        .space_after
                        .map(|twips| Value::Length(Length::from_twips(i32::from(twips)))),
                    4 => properties
                        .indent_left
                        .map(|twips| Value::Length(Length::from_twips(twips))),
                    5 => properties
                        .indent_right
                        .map(|twips| Value::Length(Length::from_twips(twips))),
                    6 => properties
                        .indent_first_line
                        .map(|twips| Value::Length(Length::from_twips(twips))),
                    7 => Some(Value::Bool(properties.keep_with_next)),
                    8 => Some(Value::Bool(properties.keep_on_page)),
                    9 => Some(Value::Bool(properties.page_break_before)),
                    16 | 17 => paragraph_value(doc, paragraph, K::StyleId),
                    18 => paragraph_value(doc, paragraph, K::StyleName),
                    19 => paragraph_value(doc, paragraph, K::NumberFormat),
                    20 => paragraph_value(doc, paragraph, K::ListStyle),
                    21 => paragraph_value(doc, paragraph, K::Start),
                    22 => paragraph_value(doc, paragraph, K::Level),
                    _ => None,
                }
            },
            Self::Run { run, doc } => match index {
                1 => run.font_index().and_then(|ftc| {
                    doc.reader()
                        .ok()?
                        .font_name(ftc)
                        .ok()
                        .flatten()
                        .map(|name| Value::Text(Cow::Borrowed(name)))
                }),
                2 => run
                    .font_size()
                    .map(|size| Value::Length(Length::from_half_points(i32::from(size)))),
                3 => run.bold().map(Value::Bool),
                4 => run.italic().map(Value::Bool),
                5 => run
                    .color()
                    .map(|(red, green, blue)| Value::Color(Color::rgb(red, green, blue))),
                6 => run
                    .underline()
                    .map(|_| Value::Token(spec::underline_name(run.underline_style()))),
                7 => run.strikethrough().map(Value::Bool),
                8 => run.all_caps().map(Value::Bool),
                9 => run.small_caps().map(Value::Bool),
                10 => run.is_superscript().then_some(Value::Bool(true)),
                11 => run.is_subscript().then_some(Value::Bool(true)),
                12 => super::parts::run_style(doc, run),
                _ => None,
            },
            Self::Table(table) => match index {
                0 => table
                    .row_count()
                    .ok()
                    .and_then(|count| i64::try_from(count).ok())
                    .map(Value::Int),
                1 => table
                    .column_count()
                    .ok()
                    .and_then(|count| i64::try_from(count).ok())
                    .map(Value::Int),
                _ => None,
            },
            Self::Row(row) => match index {
                0 => Some(Value::Bool(row.is_header())),
                _ => None,
            },
            Self::Cell(cell) => match index {
                1 => cell
                    .background_color()
                    .map(|(red, green, blue)| Value::Color(Color::rgb(red, green, blue))),
                _ => None,
            },
            Self::HeaderFooter(story) => match index {
                1 => Some(Value::Token(story_type(story.header_footer_type))),
                _ => None,
            },
            Self::Note(number) => match index {
                1 => i64::try_from(number).ok().map(Value::Int),
                _ => None,
            },
            Self::Comment(comment) => match index {
                1 => Some(Value::Text(Cow::Borrowed(comment.author.as_str()))),
                2 => Some(Value::Text(Cow::Borrowed(comment.initials.as_str()))),
                _ => None,
            },
            Self::Count(first, second) => match index {
                0 => i64::try_from(first).ok().map(Value::Int),
                1 => i64::try_from(second).ok().map(Value::Int),
                _ => None,
            },
            Self::Style { doc, istd, view } => super::parts::style_value(doc, istd, view, index),
            Self::AbstractNum { doc, at } => super::parts::abstract_num_value(doc, at, index),
            Self::Num { doc, at } => super::parts::num_value(doc, at, index),
            Self::Level {
                doc,
                definition,
                level,
                view,
            } => super::parts::level_value(doc, definition, level, view, index),
            Self::Hyperlink { doc, link } => super::marks::link_value(doc, link, index),
            Self::Bookmark(bookmark) => super::marks::bookmark_value(bookmark, index),
            Self::Field(field) => super::marks::field_value(field, index),
            Self::Toc(field) => super::marks::toc_value(field, index),
            Self::Section(section) => super::marks::section_value(section, index),
            Self::Tab(stop) => super::marks::tab_value(stop, index),
            Self::TextBox(text_box) => super::marks::text_box_value(text_box, index),
            Self::Shape(anchor) => super::marks::shape_value(anchor, index),
        }
    }
}

pub(super) const fn story_type(kind: litchi_doc::HeaderFooterType) -> &'static str {
    use litchi_doc::HeaderFooterType as T;
    match kind {
        T::FirstPageHeader | T::FirstPageFooter => "first",
        T::EvenPageHeader | T::EvenPageFooter => "even",
        T::OddPageHeader | T::OddPageFooter => "default",
    }
}

impl std::fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("doc::Node")
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

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        let count = match &self.item {
            Item::Bare => match self.kind {
                Kind::Body => self.doc.body_child_count()?,
                Kind::Comments => self.doc.reader()?.comment_count(),
                _ => Children::of(self).count(),
            },
            Item::BodyParagraph { .. } | Item::Paragraph(_) => self
                .native_paragraph()?
                .map_or(Ok(0), |paragraph| paragraph.runs().map(|runs| runs.len()))?,
            Item::Run { .. } => 0,
            Item::Table(table) => table.row_count()?,
            Item::Row(row) => row.cell_count()?,
            Item::Cell(cell) => cell.paragraphs()?.len(),
            Item::Story { story, .. } => story.paragraphs().len(),
            Item::Style(..)
            | Item::Num(_)
            | Item::Level { .. }
            | Item::Hyperlink(_)
            | Item::Bookmark(_)
            | Item::Field(_)
            | Item::Toc(_)
            | Item::Section(_)
            | Item::Tab(_)
            | Item::TextBox(_)
            | Item::Shape(_) => 0,
            Item::AbstractNum(_) => Children::of(self).count(),
        };
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    fn format(&self) -> Result<Format<'_>> {
        self.format_props()
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        self.children()
    }
}
