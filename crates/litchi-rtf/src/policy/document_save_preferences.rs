/// Passive read-only recommendation represented by `readonlyrecommended`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum DocumentReadOnlyRecommendation {
    #[default]
    Unspecified,
    Recommended,
}

/// Passive thumbnail-generation preference represented by `saveprevpict`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum DocumentThumbnailPreference {
    #[default]
    Unspecified,
    RequiredIfSupported,
}

/// Passive save-related document preferences.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentSavePreferences {
    pub read_only: DocumentReadOnlyRecommendation,
    pub thumbnail: DocumentThumbnailPreference,
}

impl DocumentSavePreferences {
    pub fn is_empty(self) -> bool {
        self == Self::default()
    }
}
