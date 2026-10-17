//! Typed generic nodes for ODT documents, borrowed from the leaf body index.

use std::borrow::Cow;
use std::sync::{Arc, OnceLock};

use litchi_core::{FileFormat, Position};
use litchi_odt::body::{Body, NodeId, Style};

use super::Odt;
use super::children::Children;
use super::props::Props;
use super::spec::{self, Kind};
use crate::error::Result;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, ViewOptions,
};

/// One step of an address: a kind and the zero-based ordinal among the
/// parent's children of that kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Step {
    pub(crate) kind: Kind,
    pub(crate) ordinal: u32,
}

impl Step {
    pub(crate) const fn new(kind: Kind, ordinal: u32) -> Self {
        Self { kind, ordinal }
    }

    fn positional(self) -> PathStep<'static> {
        use crate::path::navigate::ElementKind;
        match self.kind {
            Kind::Body | Kind::Styles | Kind::Metadata | Kind::Numbering => {
                PathStep::new(self.kind.name())
            },
            kind => PathStep::at(kind.name(), Position::new(self.ordinal as usize)),
        }
    }
}

/// Steps kept inline: `/body/tbl/tr/tc/p/r` fits with room to spare;
/// deeper addresses (nested tables, sections, lists) share heap storage.
const INLINE_STEPS: usize = 8;

/// A compact positional address.
#[derive(Clone)]
pub(crate) enum Address {
    Inline {
        len: u8,
        steps: [Step; INLINE_STEPS],
    },
    Heap(Arc<[Step]>),
}

impl Address {
    pub(crate) const fn root() -> Self {
        Self::Inline {
            len: 0,
            steps: [Step::new(Kind::Document, 0); INLINE_STEPS],
        }
    }

    pub(crate) fn steps(&self) -> &[Step] {
        match self {
            Self::Inline { len, steps } => &steps[..usize::from(*len)],
            Self::Heap(steps) => steps,
        }
    }

    pub(crate) fn with(&self, step: Step) -> Self {
        match self {
            Self::Inline { len, steps } if usize::from(*len) < INLINE_STEPS => {
                let mut steps = *steps;
                steps[usize::from(*len)] = step;
                Self::Inline {
                    len: *len + 1,
                    steps,
                }
            },
            _ => {
                let mut all = Vec::with_capacity(self.steps().len() + 1);
                all.extend_from_slice(self.steps());
                all.push(step);
                Self::Heap(all.into())
            },
        }
    }
}

/// What a node stands on in the leaf.
#[derive(Clone, Copy)]
pub(crate) enum Target<'f> {
    /// The document root, `/styles` or `/metadata`.
    Synthetic,
    /// A node of the body index.
    Body(Body<'f>, NodeId),
    /// A common style.
    Style(Style<'f>),
    /// A part of record 0328, read through the handle's part index.
    Part(Part),
}

/// A part of record 0328: an index into the handle's part index
/// ([`litchi_odt::body::Parts`]), its fields or its equations.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Part {
    /// `/numbering`.
    Numbering,
    /// A header or footer region (index among every region).
    Region(u32),
    /// A list style, as a definition or an instance.
    List(u32),
    /// A level (list style, level index).
    Level(u32, u32),
    /// A body text field (index among [`litchi_odt::body::fields`]).
    Field(u32),
    /// A table of contents.
    Toc(u32),
    /// A tab stop (paragraph, index among its stops).
    Tab(NodeId, u32),
    /// A drawing shape.
    Shape(u32),
    /// An equation (index among the handle's equations).
    Equation(u32),
    /// A table column (table, column ordinal).
    Column(NodeId, u32),
}

// Record 0287: nodes move through every query; keep this one no larger
// than it is (the unified node is as large as its largest format node).
#[cfg(target_pointer_width = "64")]
const _: () = assert!(size_of::<Node<'static>>() <= 176);

/// A node of an ODT document: the root, the body and every element the leaf
/// body index models, the common styles and the document properties.
#[derive(Clone)]
pub struct Node<'f> {
    pub(super) odt: &'f Odt,
    pub(super) kind: Kind,
    pub(super) target: Target<'f>,
    pub(super) address: Address,
    /// Text the node spells out on first use (record 0287 keeps it in one
    /// slot so the node stays small).
    spelled: OnceLock<Spelled>,
}

/// What a node spells out on first use.
#[derive(Clone)]
enum Spelled {
    /// A path deeper than [`NodePath`] carries by value.
    Steps(Box<[PathStep<'static>]>),
    /// A style name that the leaf had to unescape.
    Name(Box<str>),
}

impl<'f> Node<'f> {
    pub(super) fn new(odt: &'f Odt, kind: Kind, target: Target<'f>, address: Address) -> Self {
        Self {
            odt,
            kind,
            target,
            address,
            spelled: OnceLock::new(),
        }
    }

    pub(super) fn style(odt: &'f Odt, style: Style<'f>, ordinal: u32) -> Self {
        Self::new(
            odt,
            Kind::Style,
            Target::Style(style),
            Address::root()
                .with(Step::new(Kind::Styles, 0))
                .with(Step::new(Kind::Style, ordinal)),
        )
    }

    /// The style name of a style node, the step its path names.
    pub(super) fn style_name(&self) -> Option<&str> {
        let (Kind::Style, Target::Style(style)) = (self.kind, self.target) else {
            return None;
        };
        match style.name()? {
            Cow::Borrowed(name) => Some(name),
            Cow::Owned(name) => match self.spelled.get_or_init(|| Spelled::Name(name.into())) {
                Spelled::Name(name) => Some(name),
                Spelled::Steps(_) => None,
            },
        }
    }

    /// The element kind.
    #[must_use]
    pub const fn element_kind(&self) -> Kind {
        self.kind
    }

    /// The leaf body node, for body-backed kinds.
    #[must_use]
    pub fn leaf(&self) -> Option<litchi_odt::body::Node<'f>> {
        match self.target {
            Target::Body(body, id) => body.node(id),
            _ => None,
        }
    }

    /// The part a part node stands on.
    pub(crate) fn part(&self) -> Option<Part> {
        match self.target {
            Target::Part(part) => Some(part),
            _ => None,
        }
    }

    pub(crate) fn id(&self) -> Option<NodeId> {
        match self.target {
            Target::Body(_, id) => Some(id),
            _ => None,
        }
    }

    pub(crate) fn address(&self) -> &Address {
        &self.address
    }

    /// The node's visible text, for the kinds that hold text.
    ///
    /// # Errors
    ///
    /// Returns `corrupt_file` when the node's markup cannot be read.
    pub fn text(&self) -> Result<Option<Cow<'f, str>>> {
        if let Target::Part(part) = self.target {
            return super::parts::text(self.odt, part);
        }
        let Some(leaf) = self.leaf() else {
            return Ok(None);
        };
        match self.kind {
            Kind::Paragraph
            | Kind::Heading
            | Kind::Run
            | Kind::Hyperlink
            | Kind::ListItem
            | Kind::Cell
            | Kind::TextBox
            | Kind::Footnote
            | Kind::Endnote
            | Kind::Comment => leaf.text().map(Some).map_err(super::core_error),
            _ => Ok(None),
        }
    }

    /// The OfficeCLI type name.
    #[must_use]
    pub fn kind(&self) -> &'static str {
        self.kind.type_name()
    }

    #[must_use]
    pub fn file_format(&self) -> FileFormat {
        self.odt.format()
    }

    /// The canonical positional path.
    #[must_use]
    pub fn path(&self) -> NodePath<'_> {
        if let Some(name) = self.style_name() {
            if let Some(path) = NodePath::inline(&[PathStep::new("styles"), PathStep::new(name)]) {
                return path;
            }
        }
        let steps = self.address.steps();
        if let Some(path) = NodePath::inline_with(steps.len(), |index| steps[index].positional()) {
            return path;
        }
        match self.spelled.get_or_init(|| {
            Spelled::Steps(
                self.address
                    .steps()
                    .iter()
                    .map(|step| step.positional())
                    .collect(),
            )
        }) {
            Spelled::Steps(steps) => NodePath::new(steps),
            Spelled::Name(_) => NodePath::ROOT,
        }
    }

    /// The positional path; the same as [`Self::path`].
    #[must_use]
    pub fn positional_path(&self) -> NodePath<'_> {
        self.path()
    }

    /// The node's children, in document order: the structural kinds (the
    /// parts of record 0328 under a structural parent are reached by path
    /// and query only).
    pub fn children(&self) -> ViewChildren<'_, 'f> {
        match Children::view(self) {
            Ok(children) => ViewChildren::odt(children),
            Err(error) => ViewChildren::failed(error),
        }
    }

    /// The common style a paragraph, heading or run is based on.
    fn common_style(&self) -> Option<Cow<'f, str>> {
        let leaf = self.leaf()?;
        match self.kind {
            Kind::Paragraph | Kind::Heading | Kind::Run => leaf.common_style_name(),
            _ => None,
        }
    }
}

impl std::fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("odt::Node")
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        self.kind()
    }

    fn file_format(&self) -> FileFormat {
        self.odt.format()
    }

    fn path(&self) -> NodePath<'_> {
        self.path()
    }

    fn positional_path(&self) -> NodePath<'_> {
        self.path()
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        Node::text(self)
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        if let Target::Style(style) = self.target {
            return Ok(style.display_name().map(Preview::Text));
        }
        Ok(Node::text(self)?.map(Preview::Prefix))
    }

    fn style(&self) -> Result<Option<Cow<'_, str>>> {
        Ok(match self.kind {
            Kind::Paragraph | Kind::Heading => self.common_style(),
            _ => None,
        })
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        let count = match (self.kind, self.target) {
            (Kind::Document, _) => 3,
            (Kind::Styles, _) => self.odt.body()?.common_styles().count(),
            (_, Target::Body(..) | Target::Part(_)) => Children::view(self)?.count(),
            _ => 0,
        };
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    fn format(&self) -> Result<Format<'_>> {
        let props = match self.kind {
            Kind::Paragraph => spec::PARAGRAPH_PROPS,
            Kind::Heading => spec::HEADING_PROPS,
            Kind::Run => spec::RUN_PROPS,
            Kind::Hyperlink => spec::HYPERLINK_PROPS,
            Kind::List => spec::LIST_PROPS,
            Kind::ListItem => spec::LIST_ITEM_PROPS,
            Kind::Table => spec::TABLE_PROPS,
            Kind::Row => spec::ROW_PROPS,
            Kind::Cell => spec::CELL_PROPS,
            Kind::Section => spec::SECTION_PROPS,
            Kind::Frame => spec::FRAME_PROPS,
            Kind::Picture => spec::PICTURE_PROPS,
            Kind::TextBox => spec::TEXT_BOX_PROPS,
            Kind::Footnote | Kind::Endnote => spec::NOTE_PROPS,
            Kind::Comment => spec::COMMENT_PROPS,
            Kind::Bookmark => spec::BOOKMARK_PROPS,
            Kind::Style => spec::STYLE_PROPS,
            Kind::Metadata => spec::METADATA_PROPS,
            Kind::Header | Kind::Footer => spec::STORY_PROPS,
            Kind::Numbering => spec::NUMBERING_PROPS,
            Kind::AbstractNum => spec::ABSTRACT_NUM_PROPS,
            Kind::Num => spec::NUM_PROPS,
            Kind::Level => spec::LEVEL_PROPS,
            Kind::Field => spec::FIELD_PROPS,
            Kind::Toc => spec::TOC_PROPS,
            Kind::Tab => spec::TAB_PROPS,
            Kind::Shape => spec::SHAPE_PROPS,
            Kind::Equation => spec::EQUATION_PROPS,
            Kind::Column => spec::COLUMN_PROPS,
            Kind::Document | Kind::Body | Kind::Styles => return Ok(Format::empty()),
        };
        let source = Props::of(self)?;
        Ok(Format::new(props, Source::Odt(source)).without(&["text"]))
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        self.children()
    }
}

const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Node<'static>>();
    assert!(size_of::<Node<'static>>() <= 256);
};
