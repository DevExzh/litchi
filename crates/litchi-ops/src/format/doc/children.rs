//! Lazy root/body children for the DOC node adapter.

use super::Doc;
use super::node::Node;
use crate::error::Result;
use litchi_core::Position;

/// Children of a DOC root or body.
pub(crate) struct Children<'f> {
    pub(super) doc: &'f Doc,
    pub(super) next: usize,
    pub(super) body: bool,
    pub(super) done: bool,
    /// The native body traversal owns the scan cursor.  Nodes yielded from it
    /// carry the borrowed text so query predicates do not rescan from the
    /// beginning for every paragraph.
    pub(super) paragraphs: Option<litchi_doc::BodyParagraphs<'f>>,
    pub(super) error: Option<crate::Error>,
}

impl<'f> Children<'f> {
    pub(super) fn of(parent: &Node<'f>) -> Self {
        let body = matches!(parent.kind, super::spec::Kind::Body);
        let (paragraphs, error) = if body {
            match parent.doc.body_paragraphs() {
                Ok(paragraphs) => (Some(paragraphs), None),
                Err(error) => (None, Some(error)),
            }
        } else {
            (None, None)
        };
        Self {
            doc: parent.doc,
            next: 0,
            body,
            done: matches!(parent.kind, super::spec::Kind::Paragraph),
            paragraphs,
            error,
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        if !self.body {
            self.done = true;
            return Some(Ok(Node::body(self.doc)));
        }
        if let Some(error) = self.error.take() {
            self.done = true;
            return Some(Err(error));
        }
        let Some(paragraphs) = self.paragraphs.as_mut() else {
            self.done = true;
            return Some(Err(crate::Error::new(
                crate::Code::InternalError,
                "the DOC body paragraph iterator was not initialized",
            )
            .with_format(litchi_core::FileFormat::Doc)));
        };
        match paragraphs.next() {
            Some(Ok(text)) => {
                let position = Position::new(self.next);
                self.next = self.next.saturating_add(1);
                Some(Ok(Node::paragraph_with_text(self.doc, position, text)))
            },
            None => {
                self.done = true;
                None
            },
            Some(Err(error)) => {
                self.done = true;
                Some(Err(error.into()))
            },
        }
    }
}
