//! Typed resource dimensions of the primary presentation transaction.

/// One bounded dimension of [`super::Limits`].
///
/// [`crate::Error::EditLimit`] reports which dimension refused an edit,
/// durable patch or history operation. Variants carry structural metadata
/// only, never presentation content.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Changed parts in one patch ([`super::Limits::max_parts`]).
    Parts,
    /// Encoded durable-patch bytes ([`super::Limits::max_patch_bytes`]).
    PatchBytes,
    /// Replacement text bytes ([`super::Limits::max_text_bytes`]).
    TextBytes,
    /// Retained history entries ([`super::Limits::max_history_entries`]).
    HistoryEntries,
    /// Aggregate encoded history bytes ([`super::Limits::max_history_bytes`]).
    HistoryBytes,
    /// Retained candidate bytes ([`super::Limits::max_retained_candidate_bytes`]).
    RetainedCandidateBytes,
    /// Retained markup-compatibility bytes.
    RetainedMceBytes,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::Parts => "changed parts",
            Self::PatchBytes => "durable patch bytes",
            Self::TextBytes => "replacement text bytes",
            Self::HistoryEntries => "history entries",
            Self::HistoryBytes => "history bytes",
            Self::RetainedCandidateBytes => "retained candidate bytes",
            Self::RetainedMceBytes => "retained markup-compatibility bytes",
        })
    }
}
