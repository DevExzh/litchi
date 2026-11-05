//! The children of an object-tree node, filtered by kind.

use super::OdfTree;
use super::node::Node;
use crate::error::Result;

/// The children of one item whose kind ordinals are in a mask.
pub(crate) struct Children<'f> {
    file: &'f OdfTree,
    children: std::slice::Iter<'f, u32>,
    mask: u128,
}

impl<'f> Children<'f> {
    pub(crate) fn filtered(file: &'f OdfTree, parent: u32, mask: u128) -> Self {
        Self {
            file,
            children: file.tree().item(parent).children().iter(),
            mask,
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        let tree = self.file.tree();
        self.children
            .by_ref()
            .find(|child| self.mask & (1_u128 << (tree.item(**child).kind() & 127)) != 0)
            .map(|child| Ok(Node::new(self.file, *child)))
    }
}

impl std::iter::FusedIterator for Children<'_> {}
