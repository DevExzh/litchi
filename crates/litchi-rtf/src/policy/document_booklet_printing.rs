/// Passive document-level booklet-printing metadata.
///
/// This crate preserves these requests but never paginates, imposes, prints,
/// changes orientation, or enables duplex behavior because of them.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentBookletPrinting {
    /// `\bookfold`: booklet printing was requested.
    pub book_fold: bool,
    /// `\bookfoldrev`: reverse booklet printing was requested.
    pub reverse_book_fold: bool,
    /// Explicit `\bookfoldsheetsN`, preserved separately from omission.
    ///
    /// RTF 1.9.1 requires the value to be a nonnegative multiple of four.
    /// Explicit zero is retained because it is emitted by real producers.
    pub sheets_per_booklet: Option<u32>,
}

impl DocumentBookletPrinting {
    /// Return whether no booklet-printing metadata was present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.book_fold && !self.reverse_book_fold && self.sheets_per_booklet.is_none()
    }
}
