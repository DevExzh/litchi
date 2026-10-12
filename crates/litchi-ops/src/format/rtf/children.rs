//! Lazy traversal of native RTF body paragraphs.
use super::{Rtf, node::Node, spec::Kind};
use crate::Result;
use litchi_core::Position;

pub(crate) struct Children<'f> {
    pub(super) doc: &'f Rtf,
    pub(super) next: usize,
    pub(super) body: bool,
    pub(super) done: bool,
    pub(super) paragraphs: Option<litchi_rtf::text::Paragraphs<'f>>,
}
impl<'f> Children<'f> {
    pub(super) fn of(parent: &Node<'f>) -> Self {
        let body = parent.kind == Kind::Body;
        Self {
            doc: parent.doc,
            next: 0,
            body,
            done: parent.kind == Kind::Paragraph,
            paragraphs: body.then(|| parent.doc.0.package.body().paragraphs()),
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
        let paragraph = self.paragraphs.as_mut()?.next();
        match paragraph {
            Some(paragraph) => {
                let position = Position::new(self.next);
                self.next = self.next.saturating_add(1);
                Some(Ok(Node::paragraph(self.doc, position, paragraph)))
            },
            None => {
                self.done = true;
                None
            },
        }
    }
}
