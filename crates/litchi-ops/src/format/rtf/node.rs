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

/// A root, body or ordinary main-story paragraph in a RTF document.
#[derive(Clone, Copy)]
pub struct Node<'f> {
    pub(super) doc: &'f Rtf,
    pub(super) kind: Kind,
    pub(super) position: Option<usize>,
    /// Borrowed native paragraph retained by both direct lookup and lazy walks.
    pub(super) paragraph: Option<litchi_rtf::text::Paragraph<'f>>,
    pub(super) steps: [PathStep<'static>; 2],
}

impl<'f> Node<'f> {
    pub(super) const fn root(doc: &'f Rtf) -> Self {
        Self {
            doc,
            kind: Kind::Document,
            position: None,
            paragraph: None,
            steps: [PathStep::new("document"), PathStep::new("body")],
        }
    }

    pub(super) const fn body(doc: &'f Rtf) -> Self {
        Self {
            doc,
            kind: Kind::Body,
            position: None,
            paragraph: None,
            steps: [PathStep::new("body"), PathStep::new("p")],
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
            position: Some(position.get()),
            paragraph: Some(paragraph),
            steps: [PathStep::new("body"), PathStep::at("p", position)],
        }
    }

    /// The element kind.
    #[must_use]
    pub const fn element_kind(&self) -> Kind {
        self.kind
    }

    /// The paragraph's zero-based main-story position.
    #[must_use]
    pub const fn paragraph_position(&self) -> Option<Position> {
        match self.position {
            Some(position) => Some(Position::new(position)),
            None => None,
        }
    }

    /// The paragraph text for a paragraph node.
    pub fn paragraph_text(&self) -> Result<Option<Cow<'f, str>>> {
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
            Kind::Paragraph => NodePath::new(&self.steps),
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
        self.paragraph_text()
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        Ok(self.paragraph_text()?.map(Preview::Prefix))
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        let count = match self.kind {
            Kind::Document => 1,
            Kind::Body => self.doc.0.package.paragraph_count(),
            Kind::Paragraph => 0,
        };
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    fn format(&self) -> Result<Format<'_>> {
        Ok(self.paragraph.map_or_else(Format::empty, |paragraph| {
            Format::new(
                super::spec::PARAGRAPH_PROPS,
                crate::node::Source::Rtf(paragraph),
            )
            .without(&["text"])
        }))
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        self.children()
    }
}
