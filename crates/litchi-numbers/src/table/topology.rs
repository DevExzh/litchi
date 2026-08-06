//! Section-relative table topology edits.
//!
//! These values describe semantic table coordinates only. The concrete IWA
//! adapter resolves them against native header/body/footer counts and performs
//! the graph, formula, merge, UID, and wire updates transactionally.

/// A row insertion whose index is relative to a semantic table section.
///
/// Section-relative coordinates remove the ambiguity at the boundary between
/// body and footer rows. An index equal to the selected section's length
/// appends to that section.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum RowInsertion {
    /// Insert into the header-row section.
    Header { index: usize },
    /// Insert into the ordinary body-row section.
    Body { index: usize },
    /// Insert into the footer-row section.
    Footer { index: usize },
}

impl RowInsertion {
    /// Insert at a section-relative header-row index.
    #[must_use]
    pub const fn header(index: usize) -> Self {
        Self::Header { index }
    }

    /// Insert at a section-relative body-row index.
    #[must_use]
    pub const fn body(index: usize) -> Self {
        Self::Body { index }
    }

    /// Insert at a section-relative footer-row index.
    #[must_use]
    pub const fn footer(index: usize) -> Self {
        Self::Footer { index }
    }
}

/// A column insertion whose index is relative to a semantic table section.
///
/// An index equal to the selected section's length appends to that section.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ColumnInsertion {
    /// Insert into the header-column section.
    Header { index: usize },
    /// Insert into the ordinary body-column section.
    Body { index: usize },
}

impl ColumnInsertion {
    /// Insert at a section-relative header-column index.
    #[must_use]
    pub const fn header(index: usize) -> Self {
        Self::Header { index }
    }

    /// Insert at a section-relative body-column index.
    #[must_use]
    pub const fn body(index: usize) -> Self {
        Self::Body { index }
    }
}

/// A row deletion whose index is relative to a semantic table section.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum RowDeletion {
    /// Delete from the header-row section.
    Header { index: usize },
    /// Delete from the ordinary body-row section.
    Body { index: usize },
    /// Delete from the footer-row section.
    Footer { index: usize },
}

impl RowDeletion {
    /// Delete the row at a section-relative header-row index.
    #[must_use]
    pub const fn header(index: usize) -> Self {
        Self::Header { index }
    }

    /// Delete the row at a section-relative body-row index.
    #[must_use]
    pub const fn body(index: usize) -> Self {
        Self::Body { index }
    }

    /// Delete the row at a section-relative footer-row index.
    #[must_use]
    pub const fn footer(index: usize) -> Self {
        Self::Footer { index }
    }
}

/// A column deletion whose index is relative to a semantic table section.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ColumnDeletion {
    /// Delete from the header-column section.
    Header { index: usize },
    /// Delete from the ordinary body-column section.
    Body { index: usize },
}

impl ColumnDeletion {
    /// Delete the column at a section-relative header-column index.
    #[must_use]
    pub const fn header(index: usize) -> Self {
        Self::Header { index }
    }

    /// Delete the column at a section-relative body-column index.
    #[must_use]
    pub const fn body(index: usize) -> Self {
        Self::Body { index }
    }
}

#[cfg(test)]
mod tests {
    use super::{ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion};

    #[test]
    fn constructors_preserve_section_relative_coordinates() {
        assert_eq!(RowInsertion::header(2), RowInsertion::Header { index: 2 });
        assert_eq!(RowInsertion::body(3), RowInsertion::Body { index: 3 });
        assert_eq!(RowInsertion::footer(4), RowInsertion::Footer { index: 4 });
        assert_eq!(ColumnInsertion::header(5), ColumnInsertion::Header { index: 5 });
        assert_eq!(ColumnInsertion::body(6), ColumnInsertion::Body { index: 6 });
        assert_eq!(RowDeletion::header(2), RowDeletion::Header { index: 2 });
        assert_eq!(RowDeletion::body(3), RowDeletion::Body { index: 3 });
        assert_eq!(RowDeletion::footer(4), RowDeletion::Footer { index: 4 });
        assert_eq!(ColumnDeletion::header(5), ColumnDeletion::Header { index: 5 });
        assert_eq!(ColumnDeletion::body(6), ColumnDeletion::Body { index: 6 });
    }
}
