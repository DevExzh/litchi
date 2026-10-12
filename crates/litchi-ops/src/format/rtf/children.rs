//! Lazy traversal of native RTF body paragraphs and paragraph runs.

use super::{Rtf, node::Node, spec::Kind};
use crate::Result;
use litchi_core::Position;

/// A lazy child iterator over the native RTF story. Paragraph and run nodes
/// retain the native iterators rather than building a second tree.
pub(crate) struct Children<'f> {
    pub(super) doc: &'f Rtf,
    pub(super) next: usize,
    pub(super) body: bool,
    pub(super) done: bool,
    pub(super) paragraphs: Option<litchi_rtf::text::Paragraphs<'f>>,
    pub(super) paragraph_position: Option<Position>,
    pub(super) paragraph: Option<litchi_rtf::text::Paragraph<'f>>,
    pub(super) runs: Option<litchi_rtf::text::Runs<'f>>,
}

impl<'f> Children<'f> {
    pub(super) const fn empty(doc: &'f Rtf) -> Self {
        Self {
            doc,
            next: 0,
            body: false,
            done: true,
            paragraphs: None,
            paragraph_position: None,
            paragraph: None,
            runs: None,
        }
    }

    pub(super) fn of(parent: &Node<'f>) -> Self {
        let body = parent.kind == Kind::Body;
        let paragraph = parent.kind == Kind::Paragraph;
        Self {
            doc: parent.doc,
            next: 0,
            body,
            done: parent.kind == Kind::Run,
            paragraphs: body.then(|| parent.doc.0.package.body().paragraphs()),
            paragraph_position: paragraph.then_some(parent.paragraph_position).flatten(),
            paragraph: paragraph.then_some(parent.paragraph).flatten(),
            runs: paragraph
                .then_some(parent.paragraph)
                .flatten()
                .map(|value| value.runs()),
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        if self.body {
            let paragraph = self.paragraphs.as_mut()?.next();
            return match paragraph {
                Some(paragraph) => {
                    let position = Position::new(self.next);
                    self.next = self.next.saturating_add(1);
                    Some(Ok(Node::paragraph(self.doc, position, paragraph)))
                },
                None => {
                    self.done = true;
                    None
                },
            };
        }
        if let Some(runs) = self.runs.as_mut() {
            let Some(run) = runs.next() else {
                self.done = true;
                return None;
            };
            let Some(paragraph_position) = self.paragraph_position else {
                self.done = true;
                return None;
            };
            let Some(paragraph) = self.paragraph else {
                self.done = true;
                return None;
            };
            let position = Position::new(self.next);
            self.next = self.next.saturating_add(1);
            return Some(Ok(Node::run(
                self.doc,
                paragraph_position,
                position,
                paragraph,
                run,
            )));
        }
        self.done = true;
        Some(Ok(Node::body(self.doc)))
    }
}
