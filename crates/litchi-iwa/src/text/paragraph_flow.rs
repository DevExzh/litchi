//! Paragraph pagination and automatic-hyphenation controls.

/// Whether a paragraph participates in automatic hyphenation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphHyphenation {
    /// Permit automatic hyphenation when it is enabled for the document.
    #[default]
    Automatic,
    /// Prevent automatic hyphenation for this paragraph.
    Prevented,
}

impl ParagraphHyphenation {
    pub(crate) const fn native_value(self) -> bool {
        matches!(self, Self::Automatic)
    }

    pub(crate) const fn from_native_value(value: bool) -> Self {
        if value {
            Self::Automatic
        } else {
            Self::Prevented
        }
    }
}

/// Effective Pagination & Breaks settings for one uniform paragraph style.
///
/// These controls correspond to Pages' Text → More inspector. The underlying
/// paragraph style is shared by Pages, Numbers, and Keynote text storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParagraphFlow {
    keep_lines_together: bool,
    keep_with_next: bool,
    start_on_new_page: bool,
    prevent_widow_orphan_lines: bool,
    hyphenation: ParagraphHyphenation,
}

impl ParagraphFlow {
    /// Construct the native effective defaults.
    pub const fn new() -> Self {
        Self {
            keep_lines_together: false,
            keep_with_next: false,
            start_on_new_page: false,
            prevent_widow_orphan_lines: true,
            hyphenation: ParagraphHyphenation::Automatic,
        }
    }

    pub const fn with_keep_lines_together(mut self, enabled: bool) -> Self {
        self.keep_lines_together = enabled;
        self
    }

    pub const fn with_keep_with_next(mut self, enabled: bool) -> Self {
        self.keep_with_next = enabled;
        self
    }

    pub const fn with_start_on_new_page(mut self, enabled: bool) -> Self {
        self.start_on_new_page = enabled;
        self
    }

    pub const fn with_prevent_widow_orphan_lines(mut self, enabled: bool) -> Self {
        self.prevent_widow_orphan_lines = enabled;
        self
    }

    pub const fn with_hyphenation(mut self, hyphenation: ParagraphHyphenation) -> Self {
        self.hyphenation = hyphenation;
        self
    }

    pub const fn keeps_lines_together(self) -> bool {
        self.keep_lines_together
    }

    pub const fn keeps_with_next(self) -> bool {
        self.keep_with_next
    }

    pub const fn starts_on_new_page(self) -> bool {
        self.start_on_new_page
    }

    pub const fn prevents_widow_orphan_lines(self) -> bool {
        self.prevent_widow_orphan_lines
    }

    pub const fn hyphenation(self) -> ParagraphHyphenation {
        self.hyphenation
    }
}

impl Default for ParagraphFlow {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_and_builders_match_the_native_inspector() {
        let defaults = ParagraphFlow::default();
        assert!(!defaults.keeps_lines_together());
        assert!(!defaults.keeps_with_next());
        assert!(!defaults.starts_on_new_page());
        assert!(defaults.prevents_widow_orphan_lines());
        assert_eq!(defaults.hyphenation(), ParagraphHyphenation::Automatic);

        let custom = defaults
            .with_keep_lines_together(true)
            .with_keep_with_next(true)
            .with_start_on_new_page(true)
            .with_prevent_widow_orphan_lines(false)
            .with_hyphenation(ParagraphHyphenation::Prevented);
        assert!(custom.keeps_lines_together());
        assert!(custom.keeps_with_next());
        assert!(custom.starts_on_new_page());
        assert!(!custom.prevents_widow_orphan_lines());
        assert_eq!(custom.hyphenation(), ParagraphHyphenation::Prevented);
    }
}
