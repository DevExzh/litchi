//! Section-relative table-axis deletion coordinates.

/// A row deletion whose index is relative to a semantic table section.
///
/// Unlike a physical row index, this type makes it explicit whether the
/// operation removes a header, body, or footer row. Deletion indices must be
/// strictly less than the selected section's current length.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum TableRowDeletion {
    /// Delete from the header-row section.
    Header { index: usize },
    /// Delete from the ordinary body-row section.
    Body { index: usize },
    /// Delete from the footer-row section.
    Footer { index: usize },
}

impl TableRowDeletion {
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
///
/// Deletion indices must be strictly less than the selected section's current
/// length.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum TableColumnDeletion {
    /// Delete from the header-column section.
    Header { index: usize },
    /// Delete from the ordinary body-column section.
    Body { index: usize },
}

impl TableColumnDeletion {
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
