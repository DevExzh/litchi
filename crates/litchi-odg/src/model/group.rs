//! Checked group-subtree ownership.

/// One group root and its flattened descendant shape positions.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Group {
    page: usize,
    shape: usize,
    descendants: Vec<usize>,
}

impl Group {
    pub(crate) const fn parsed(page: usize, shape: usize, descendants: Vec<usize>) -> Self {
        Self {
            page,
            shape,
            descendants,
        }
    }

    /// Owning page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// Root group shape position in the page's flattened shape inventory.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Complete nested descendant positions in source order.
    #[must_use]
    pub fn descendants(&self) -> &[usize] {
        &self.descendants
    }

    /// Whether the group owns one flattened descendant position.
    #[must_use]
    pub fn contains(&self, shape: usize) -> bool {
        self.descendants.binary_search(&shape).is_ok()
    }
}
