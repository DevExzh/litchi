//! The children of an iWork object-tree node, filtered by kind.

use super::Iwork;
use super::Kids;
use super::node::Node;
use crate::error::Result;

/// The children of one item whose kind ordinals are in a mask.
pub(crate) struct Children<'f> {
    file: &'f Iwork,
    kids: Kids<'f>,
    next: usize,
    mask: u128,
}

impl<'f> Children<'f> {
    pub(crate) fn filtered(file: &'f Iwork, parent: super::At, mask: u128) -> Self {
        Self {
            file,
            kids: file.kids(parent),
            next: 0,
            mask,
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        while let Some(child) = self.kids.get(self.next) {
            self.next += 1;
            if self.mask & (1_u128 << (self.file.kind_at(child) & 127)) != 0 {
                return Some(Ok(Node::new(self.file, child)));
            }
        }
        None
    }
}

impl std::iter::FusedIterator for Children<'_> {}
