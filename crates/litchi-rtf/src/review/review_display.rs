//! Passive document review-display preferences.

/// Explicit RTF preferences for hiding tracked-review information.
///
/// These flags are metadata only. The parser does not render, accept, reject, or execute
/// revisions and comments based on them.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentReviewDisplay {
    /// `donotshowmarkup`: do not show markup while reviewing.
    pub hide_markup: bool,
    /// `donotshowcomments`: do not show comments while reviewing.
    pub hide_comments: bool,
    /// `donotshowinsdel`: do not show insertions and deletions while reviewing.
    pub hide_insertions_and_deletions: bool,
}

impl DocumentReviewDisplay {
    /// Return whether no review-display suppression flag is present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.hide_markup && !self.hide_comments && !self.hide_insertions_and_deletions
    }
}
