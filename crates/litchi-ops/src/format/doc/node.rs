//! Typed generic nodes for legacy Word documents.

use std::borrow::Cow;

use litchi_core::{FileFormat, Position};

use super::Doc;
use super::children::Children;
use super::spec::Kind;
use crate::error::Result;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, ViewOptions,
};

/// A root, body or ordinary main-story paragraph in a DOC document.
#[derive(Clone, Copy)]
pub struct Node<'f> {
    pub(super) doc: &'f Doc,
    pub(super) kind: Kind,
    pub(super) position: Option<usize>,
    /// Text yielded by the native body stream, when this node came from a
    /// children/query walk.  Point lookups leave this empty and use the
    /// native ordinal accessor instead.
    pub(super) paragraph_text: Option<&'f str>,
    pub(super) steps: [PathStep<'static>; 2],
}

impl<'f> Node<'f> {
    pub(super) const fn root(doc: &'f Doc) -> Self {
        Self {
            doc,
            kind: Kind::Document,
            position: None,
            paragraph_text: None,
            steps: [PathStep::new("document"), PathStep::new("body")],
        }
    }

    pub(super) const fn body(doc: &'f Doc) -> Self {
        Self {
            doc,
            kind: Kind::Body,
            position: None,
            paragraph_text: None,
            steps: [PathStep::new("body"), PathStep::new("p")],
        }
    }

    pub(super) const fn paragraph(doc: &'f Doc, position: Position) -> Self {
        Self {
            doc,
            kind: Kind::Paragraph,
            position: Some(position.get()),
            paragraph_text: None,
            steps: [PathStep::new("body"), PathStep::at("p", position)],
        }
    }

    pub(super) fn paragraph_with_text(
        doc: &'f Doc,
        position: Position,
        paragraph_text: &'f str,
    ) -> Self {
        Self {
            doc,
            kind: Kind::Paragraph,
            position: Some(position.get()),
            paragraph_text: Some(paragraph_text),
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
    pub fn paragraph_text(&self) -> Result<Option<&str>> {
        self.doc.check_source()?;
        match (self.kind, self.position) {
            (Kind::Paragraph, Some(position)) => {
                if let Some(text) = self.paragraph_text {
                    return Ok(Some(text));
                }
                Ok(Some(
                    self.doc
                        .body_paragraph_at(Position::new(position))?
                        .ok_or_else(|| {
                            crate::Error::new(
                                crate::Code::PathNotFound,
                                "the DOC paragraph is outside the main story",
                            )
                            .with_format(FileFormat::Doc)
                        })?,
                ))
            },
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
        ViewChildren::doc(Children::of(self))
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
        Ok(self.paragraph_text()?.map(Cow::Borrowed))
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        Ok(self
            .paragraph_text()?
            .map(|text| Preview::Prefix(Cow::Borrowed(text))))
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        let count = match self.kind {
            Kind::Document => 1,
            Kind::Body => self.doc.body_paragraph_count()?,
            Kind::Paragraph => 0,
        };
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    fn format(&self) -> Result<Format<'_>> {
        Ok(Format::empty())
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        self.children()
    }
}
