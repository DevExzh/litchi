//! Section-relative table-axis insertion coordinates.

/// A row insertion whose index is relative to a semantic table section.
///
/// Section-relative coordinates remove the ambiguity at the boundary between
/// body and footer rows. An index equal to the selected section's length
/// appends to that section.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum TableRowInsertion {
    /// Insert into the header-row section.
    Header { index: usize },
    /// Insert into the ordinary body-row section.
    Body { index: usize },
    /// Insert into the footer-row section.
    Footer { index: usize },
}

impl TableRowInsertion {
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
pub enum TableColumnInsertion {
    /// Insert into the header-column section.
    Header { index: usize },
    /// Insert into the ordinary body-column section.
    Body { index: usize },
}

impl TableColumnInsertion {
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
