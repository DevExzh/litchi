//! Typed generic nodes for RTF documents.

use std::borrow::Cow;

use litchi_core::{FileFormat, Position};

use super::Rtf;
use super::children::Children;
use super::spec::Kind;
use crate::error::Result;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, ViewOptions,
};

/// A root, body, paragraph, or borrowed character run in an RTF document.
#[derive(Clone, Copy)]
pub struct Node<'f> {
    pub(super) doc: &'f Rtf,
    pub(super) kind: Kind,
    /// The containing paragraph's zero-based main-story position.
    pub(super) paragraph_position: Option<Position>,
    /// The run's zero-based position within its containing paragraph.
    pub(super) run_position: Option<Position>,
    /// Borrowed native paragraph retained by paragraph and run nodes.
    pub(super) paragraph: Option<litchi_rtf::text::Paragraph<'f>>,
    /// Borrowed native run retained by run nodes.
    pub(super) run: Option<litchi_rtf::text::Run<'f>>,
    pub(super) steps: [PathStep<'static>; 3],
}

impl<'f> Node<'f> {
    pub(super) const fn root(doc: &'f Rtf) -> Self {
        Self {
            doc,
            kind: Kind::Document,
            paragraph_position: None,
            run_position: None,
            paragraph: None,
            run: None,
            steps: [
                PathStep::new("document"),
                PathStep::new("body"),
                PathStep::new("p"),
            ],
        }
    }

    pub(super) const fn body(doc: &'f Rtf) -> Self {
        Self {
            doc,
            kind: Kind::Body,
            paragraph_position: None,
            run_position: None,
            paragraph: None,
            run: None,
            steps: [
                PathStep::new("body"),
                PathStep::new("p"),
                PathStep::new("r"),
            ],
        }
    }

    pub(super) const fn paragraph(
        doc: &'f Rtf,
        position: Position,
        paragraph: litchi_rtf::text::Paragraph<'f>,
    ) -> Self {
        Self {
            doc,
            kind: Kind::Paragraph,
            paragraph_position: Some(position),
            run_position: None,
            paragraph: Some(paragraph),
            run: None,
            steps: [
                PathStep::new("body"),
                PathStep::at("p", position),
                PathStep::new("r"),
            ],
        }
    }

    pub(super) const fn run(
        doc: &'f Rtf,
        paragraph_position: Position,
        run_position: Position,
        paragraph: litchi_rtf::text::Paragraph<'f>,
        run: litchi_rtf::text::Run<'f>,
    ) -> Self {
        Self {
            doc,
            kind: Kind::Run,
            paragraph_position: Some(paragraph_position),
            run_position: Some(run_position),
            paragraph: Some(paragraph),
            run: Some(run),
            steps: [
                PathStep::new("body"),
                PathStep::at("p", paragraph_position),
                PathStep::at("r", run_position),
            ],
        }
    }

    /// The element kind.
    #[must_use]
    pub const fn element_kind(&self) -> Kind {
        self.kind
    }

    /// The containing paragraph's zero-based main-story position.
    #[must_use]
    pub const fn paragraph_position(&self) -> Option<Position> {
        self.paragraph_position
    }

    /// The run's zero-based position within its containing paragraph.
    #[must_use]
    pub const fn run_position(&self) -> Option<Position> {
        self.run_position
    }

    /// The borrowed native run retained by this run node.
    #[must_use]
    pub const fn as_run(&self) -> Option<litchi_rtf::text::Run<'f>> {
        self.run
    }

    /// The node's visible text. Paragraph text remains borrowed when the
    /// native story stores it contiguously; fragmented paragraphs use the
    /// native formatter's owned fallback. Runs always borrow their text.
    pub fn text(&self) -> Result<Option<Cow<'f, str>>> {
        if let Some(run) = self.run {
            return Ok(Some(Cow::Borrowed(run.text())));
        }
        Ok(self.paragraph.map(|paragraph| {
            paragraph
                .as_str()
                .map_or_else(|| Cow::Owned(paragraph.to_text()), Cow::Borrowed)
        }))
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
        match self.kind {
            Kind::Document => NodePath::new(&[]),
            Kind::Body => NodePath::new(&self.steps[..1]),
            Kind::Paragraph => NodePath::new(&self.steps[..2]),
            Kind::Run => NodePath::new(&self.steps),
        }
    }

    #[must_use]
    pub fn positional_path(&self) -> NodePath<'_> {
        self.path()
    }

    pub fn children(&self) -> ViewChildren<'_, 'f> {
        ViewChildren::rtf(Children::of(self))
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

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        let count = match self.kind {
            Kind::Document => 1,
            Kind::Body => self.doc.0.package.paragraph_count(),
            Kind::Paragraph => self
                .paragraph
                .map_or(0, |paragraph| paragraph.runs().count()),
            Kind::Run => 0,
        };
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    fn format(&self) -> Result<Format<'_>> {
        match (self.kind, self.paragraph, self.run) {
            (Kind::Paragraph, Some(paragraph), _) => Ok(Format::new(
                super::spec::PARAGRAPH_PROPS,
                crate::node::Source::RtfParagraph(paragraph),
            )
            .without(&["text"])),
            (Kind::Run, _, Some(run)) => Ok(Format::new(
                super::spec::RUN_PROPS,
                crate::node::Source::RtfRun(run),
            )
            .without(&["text"])),
            _ => Ok(Format::empty()),
        }
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        self.children()
    }
}
