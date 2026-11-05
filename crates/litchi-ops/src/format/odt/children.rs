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
    /// Ordinals of the structural kinds; parts count their own.
    counters: [u32; Kind::STRUCTURAL],
    axis: Axis<'f>,
    /// The parts of record 0328 listed after the structural children,
    /// boxed so a structural listing stays as small as it was.
    tail: Option<Box<Tail<'f>>>,
}

/// The part children of a parent, listed once its structural children are.
enum Tail<'f> {
    /// The parent and the part kinds wanted, not listed yet.
    Pending(Node<'f>, KindSet<Kind>),
    Open(std::vec::IntoIter<Node<'f>>),
}

enum Axis<'f> {
    Empty,
    /// `/body`, `/styles` and `/metadata`, in that order.
    Root(u8),
    Body(Body<'f>, litchi_odt::body::Children<'f>),
    Styles(litchi_odt::body::CommonStyles<'f>),
}

impl<'f> Children<'f> {
    /// The children of `parent` a query asks for: the structural kinds in
    /// `kinds`, then the parts in `kinds`.
    pub(super) fn of(parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        Self::listing(parent, kinds, kinds)
    }

    /// The children a node view lists: every structural child, and the
    /// parts of a part (a numbering's definitions, a definition's levels),
    /// never the parts of a structural parent (record 0328 keeps the views
    /// of the root, the body, paragraphs and tables as they were).
    pub(super) fn view(parent: &Node<'f>) -> Result<Self> {
        let parts = if parent.kind.is_part() {
            KindSet::ALL
        } else {
            KindSet::EMPTY
        };
        Self::listing(parent, KindSet::ALL, parts)
    }

    fn listing(parent: &Node<'f>, kinds: KindSet<Kind>, parts: KindSet<Kind>) -> Result<Self> {
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
        let tail = super::parts::has_part_children(parent.kind, parts)
            .then(|| Box::new(Tail::Pending(parent.clone(), parts)));
        Ok(Self {
            odt: parent.odt,
            parent: parent.address().clone(),
            kinds,
            counters: [0; Kind::STRUCTURAL],
            axis,
            tail,
        })
    }

    /// The next part child, once the structural children are listed. Kept
    /// out of line so the structural listing's hot loop keeps its shape.
    #[cold]
    #[inline(never)]
    fn next_part(&mut self) -> Option<Result<Node<'f>>> {
        loop {
            let tail = self.tail.as_deref_mut()?;
            let listed = match tail {
                Tail::Open(nodes) => return nodes.next().map(Ok),
                Tail::Pending(parent, kinds) => super::parts::children(parent, *kinds),
            };
            match listed {
                Ok(nodes) => *tail = Tail::Open(nodes.into_iter()),
                Err(error) => {
                    self.tail = None;
                    return Some(Err(error));
                },
            }
        }
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
                Axis::Empty => return self.next_part(),
                Axis::Root(next) => {
                    let kind = match *next {
                        0 => Kind::Body,
                        1 => Kind::Styles,
                        2 => Kind::Metadata,
                        _ => {
                            self.axis = Axis::Empty;
                            continue;
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
                    let Some(child) = children.next() else {
                        self.axis = Axis::Empty;
                        continue;
                    };
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
                    let Some(style) = styles.next() else {
                        self.axis = Axis::Empty;
                        continue;
                    };
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
