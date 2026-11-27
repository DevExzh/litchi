/// Theme font-resolution languages declared in the RTF document header.
///
/// These identifiers are passive metadata. This crate does not resolve or
/// substitute theme fonts from them.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentThemeLanguages {
    /// Primary theme language (`themelangN`).
    pub primary: Option<crate::LanguageId>,
    /// East Asian theme language (`themelangfeN`).
    pub east_asian: Option<crate::LanguageId>,
    /// Complex-script theme language (`themelangcsN`).
    pub complex_script: Option<crate::LanguageId>,
}

impl DocumentThemeLanguages {
    /// Return whether all theme-language controls were omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.primary.is_none() && self.east_asian.is_none() && self.complex_script.is_none()
    }
}
