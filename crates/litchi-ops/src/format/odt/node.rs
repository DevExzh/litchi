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
            Kind::Body | Kind::Styles | Kind::Metadata => PathStep::new(self.kind.name()),
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
}

/// A node of an ODT document: the root, the body and every element the leaf
/// body index models, the common styles and the document properties.
#[derive(Clone)]
pub struct Node<'f> {
    pub(super) odt: &'f Odt,
    pub(super) kind: Kind,
    pub(super) target: Target<'f>,
    pub(super) address: Address,
    /// The style name of a style node, the step its path names.
    pub(super) style_name: Option<Cow<'f, str>>,
    paths: OnceLock<Box<[PathStep<'static>]>>,
}

impl<'f> Node<'f> {
    pub(super) fn new(odt: &'f Odt, kind: Kind, target: Target<'f>, address: Address) -> Self {
        Self {
            odt,
            kind,
            target,
            address,
            style_name: None,
            paths: OnceLock::new(),
        }
    }

    pub(super) fn style(odt: &'f Odt, style: Style<'f>, ordinal: u32) -> Self {
        let mut node = Self::new(
            odt,
            Kind::Style,
            Target::Style(style),
            Address::root()
                .with(Step::new(Kind::Styles, 0))
                .with(Step::new(Kind::Style, ordinal)),
        );
        node.style_name = style.name();
        node
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
    pub const fn file_format(&self) -> FileFormat {
        FileFormat::Odt
    }

    /// The canonical positional path.
    #[must_use]
    pub fn path(&self) -> NodePath<'_> {
        if let Some(name) = &self.style_name {
            if let Some(path) = NodePath::inline(&[PathStep::new("styles"), PathStep::new(name)]) {
                return path;
            }
        }
        NodePath::new(self.paths.get_or_init(|| {
            self.address
                .steps()
                .iter()
                .map(|step| step.positional())
                .collect()
        }))
    }

    /// The positional path; the same as [`Self::path`].
    #[must_use]
    pub fn positional_path(&self) -> NodePath<'_> {
        self.path()
    }

    /// The node's children, in document order.
    pub fn children(&self) -> ViewChildren<'_, 'f> {
        match Children::of(self, crate::path::navigate::KindSet::ALL) {
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
        FileFormat::Odt
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
            (_, Target::Body(..)) => {
                Children::of(self, crate::path::navigate::KindSet::ALL)?.count()
            },
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
