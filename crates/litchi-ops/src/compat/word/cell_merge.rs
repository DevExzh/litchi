//! Format-neutral table-cell merge state, moved from the facade's
//! `document::cell_merge`.
//!
//! Word documents record cell merges in two structurally different ways:
//!
//! * **Span counts.** DOCX (`w:gridSpan`) and ODF
//!   (`table:number-columns-spanned`, `table:number-rows-spanned`) store the
//!   width of a merged range on the cell that owns it. Covered cells are
//!   either absent from the row or, in ODF, recorded as
//!   `table:covered-table-cell` elements that the reader skips.
//! * **Participation roles.** Binary DOC (`TC80.fFirstMerged`/`fMerged` and
//!   `fVertRestart`/`fVertMerge`) and RTF (`\clmgf`/`\clmrg` and
//!   `\clvmgf`/`\clvmrg`) keep every covered cell in the row and tag each one
//!   with the role it plays in the merge.
//!
//! [`CellMerge`] is the common denominator: it answers "does this cell take
//! part in a merge along this axis, and does it own the range?" for every
//! supported format. Callers that need a column count use
//! [`Row::grid_span_at`](super::Row::grid_span_at), which resolves the
//! role-based formats with the surrounding row.

/// How a table cell participates in a merged range along one axis.
///
/// The variants are ordered from "not merged" to "covered by an earlier
/// cell", and are produced identically for DOC, DOCX, RTF and ODF sources.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Hash)]
#[non_exhaustive]
pub enum CellMerge {
    /// The cell stands alone on this axis.
    #[default]
    None,
    /// The cell begins a merged range and owns the range's content.
    Start,
    /// The cell is covered by a merged range that an earlier cell began.
    ///
    /// Formats that store span counts rather than roles never report this
    /// variant, because their covered cells are not surfaced as cells at all.
    Continuation,
}

impl CellMerge {
    /// Whether the cell takes part in a merge on this axis at all.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// ```
    /// use litchi_ops::compat::word::CellMerge;
    ///
    /// assert!(CellMerge::Start.is_merged());
    /// assert!(!CellMerge::None.is_merged());
    /// ```
    #[must_use]
    pub const fn is_merged(self) -> bool {
        !matches!(self, Self::None)
    }

    /// Whether the cell owns the merged range and therefore holds its
    /// content. Unmerged cells own themselves, so this is true for both
    /// [`CellMerge::None`] and [`CellMerge::Start`]; use it to skip covered
    /// cells when reading text.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// ```
    /// use litchi_ops::compat::word::CellMerge;
    ///
    /// assert!(CellMerge::None.owns_content());
    /// assert!(!CellMerge::Continuation.owns_content());
    /// ```
    #[must_use]
    pub const fn owns_content(self) -> bool {
        !matches!(self, Self::Continuation)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_unmerged() {
        assert_eq!(CellMerge::default(), CellMerge::None);
        assert!(!CellMerge::None.is_merged());
    }

    #[test]
    fn start_is_merged_and_owns_its_content() {
        assert!(CellMerge::Start.is_merged());
        assert!(CellMerge::Start.owns_content());
    }

    #[test]
    fn continuation_is_merged_but_covered() {
        assert!(CellMerge::Continuation.is_merged());
        assert!(!CellMerge::Continuation.owns_content());
    }

    #[test]
    fn unmerged_cells_own_their_content() {
        assert!(CellMerge::None.owns_content());
    }
}
