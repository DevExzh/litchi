//! Strict list presets shared by Pages, Numbers, and Keynote.

/// A canonical paragraph-list presentation understood by all three iWork apps.
///
/// The presets describe the complete nine-level native list style rather than
/// exposing unvalidated protobuf integers or partial per-level state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphList {
    /// Ordinary paragraphs without labels.
    #[default]
    None,
    /// Apple’s standard bullet preset using the `•` marker.
    Bullet,
    /// Apple’s standard decimal-number preset.
    Numbered,
}

impl ParagraphList {
    pub(crate) const fn native_name(self) -> &'static str {
        match self {
            Self::None => "None",
            Self::Bullet => "Bullet",
            Self::Numbered => "Numbered",
        }
    }
}
