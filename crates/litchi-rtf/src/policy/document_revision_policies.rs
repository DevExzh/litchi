/// Passive revision-tracking policy controls from the RTF header.
///
/// These values are retained independently from revision content. This crate
/// does not enable tracking, infer moves, create revisions, or apply changes.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentRevisionPolicies {
    /// Explicit `trackmovesN` policy.
    pub track_moves: Option<bool>,
    /// Explicit `trackformattingN` policy.
    pub track_formatting: Option<bool>,
}

impl DocumentRevisionPolicies {
    /// Return whether both revision-policy controls were omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.track_moves.is_none() && self.track_formatting.is_none()
    }
}
