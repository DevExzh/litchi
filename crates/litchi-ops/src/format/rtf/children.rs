//! Lazy traversal of native RTF body paragraphs and paragraph runs.

use super::{Rtf, node::Node, spec::Kind};
use crate::Result;
use litchi_core::Position;

/// Retain only the native cursor for the selected axis, without a second tree.
pub(crate) struct Children<'f> {
    doc: &'f Rtf,
    next: usize,
    axis: Axis<'f>,
}

enum Axis<'f> {
    Empty,
    Body,
    Paragraphs(litchi_rtf::text::Paragraphs<'f>),
    Runs {
        paragraph: Position,
        runs: litchi_rtf::text::Runs<'f>,
    },
}

impl<'f> Children<'f> {
    pub(super) const fn empty(doc: &'f Rtf) -> Self {
        Self {
            doc,
            next: 0,
            axis: Axis::Empty,
        }
    }

    pub(super) fn of(parent: &Node<'f>) -> Self {
        let axis = match parent.kind {
            Kind::Document => Axis::Body,
            Kind::Body => Axis::Paragraphs(parent.doc.0.package.body().paragraphs()),
            Kind::Paragraph => match (parent.paragraph_position, parent.paragraph) {
                (Some(position), Some(paragraph)) => Axis::Runs {
                    paragraph: position,
                    runs: paragraph.runs(),
                },
                _ => Axis::Empty,
            },
            Kind::Run => Axis::Empty,
        };
        Self {
            doc: parent.doc,
            next: 0,
            axis,
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        let position = Position::new(self.next);
        let node = match &mut self.axis {
            Axis::Empty => return None,
            Axis::Body => {
                self.axis = Axis::Empty;
                Node::body(self.doc)
            },
            Axis::Paragraphs(paragraphs) => Node::paragraph(self.doc, position, paragraphs.next()?),
            Axis::Runs { paragraph, runs } => {
                Node::run(self.doc, *paragraph, position, runs.next()?)
            },
        };
        self.next = self.next.saturating_add(1);
        Some(Ok(node))
    }
}
