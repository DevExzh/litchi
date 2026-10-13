//! Lazy traversal of the ODT body index, the common styles and the root.

use litchi_odt::body::Body;

use super::Odt;
use super::node::{Address, Node, Step, Target};
use super::spec::Kind;
use crate::Result;
use crate::path::navigate::KindSet;

/// The children of one node whose kinds are in a set, with each child's
/// ordinal among its siblings of the same kind.
pub(crate) struct Children<'f> {
    odt: &'f Odt,
    parent: Address,
    kinds: KindSet<Kind>,
    counters: [u32; Kind::COUNT],
    axis: Axis<'f>,
}

enum Axis<'f> {
    Empty,
    /// `/body`, `/styles` and `/metadata`, in that order.
    Root(u8),
    Body(Body<'f>, litchi_odt::body::Children<'f>),
    Styles(litchi_odt::body::CommonStyles<'f>),
}

impl<'f> Children<'f> {
    pub(super) fn of(parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        let axis = match (parent.kind, parent.target) {
            (Kind::Document, _) => Axis::Root(0),
            (Kind::Styles, _) => {
                let body = parent.odt.body()?;
                Axis::Styles(body.common_styles())
            },
            (_, Target::Body(body, id)) => match body.node(id) {
                Some(node) => Axis::Body(body, node.children()),
                None => Axis::Empty,
            },
            _ => Axis::Empty,
        };
        Ok(Self {
            odt: parent.odt,
            parent: parent.address().clone(),
            kinds,
            counters: [0; Kind::COUNT],
            axis,
        })
    }

    fn count(&mut self, kind: Kind) -> u32 {
        let counter = &mut self.counters[kind as usize];
        let ordinal = *counter;
        *counter = counter.saturating_add(1);
        ordinal
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            match &mut self.axis {
                Axis::Empty => return None,
                Axis::Root(next) => {
                    let kind = match *next {
                        0 => Kind::Body,
                        1 => Kind::Styles,
                        2 => Kind::Metadata,
                        _ => {
                            self.axis = Axis::Empty;
                            return None;
                        },
                    };
                    *next += 1;
                    if !self.kinds.contains(kind) {
                        continue;
                    }
                    let address = self.parent.with(Step::new(kind, 0));
                    if kind == Kind::Body {
                        let body = match self.odt.body() {
                            Ok(body) => body,
                            Err(error) => {
                                self.axis = Axis::Empty;
                                return Some(Err(error));
                            },
                        };
                        return Some(Ok(Node::new(
                            self.odt,
                            kind,
                            Target::Body(body, litchi_odt::body::NodeId::BODY),
                            address,
                        )));
                    }
                    return Some(Ok(Node::new(self.odt, kind, Target::Synthetic, address)));
                },
                Axis::Body(body, children) => {
                    let body = *body;
                    let child = children.next()?;
                    let Some(kind) = Kind::of(child.kind()) else {
                        continue;
                    };
                    let ordinal = self.count(kind);
                    if !self.kinds.contains(kind) {
                        continue;
                    }
                    let address = self.parent.with(Step::new(kind, ordinal));
                    return Some(Ok(Node::new(
                        self.odt,
                        kind,
                        Target::Body(body, child.id()),
                        address,
                    )));
                },
                Axis::Styles(styles) => {
                    let style = styles.next()?;
                    let ordinal = self.count(Kind::Style);
                    if !self.kinds.contains(Kind::Style) {
                        continue;
                    }
                    return Some(Ok(Node::style(self.odt, style, ordinal)));
                },
            }
        }
    }
}
