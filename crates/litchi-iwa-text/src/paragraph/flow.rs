//! Archive-free paragraph pagination and hyphenation values.

/// Whether a paragraph participates in automatic hyphenation.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Hyphenation {
    /// Permit automatic hyphenation when it is enabled for the document.
    #[default]
    Automatic,
    /// Prevent automatic hyphenation for this paragraph.
    Prevented,
}

/// Effective pagination and break settings for one uniform paragraph style.
///
/// These controls correspond to Pages' Text → More inspector. The underlying
/// paragraph style is shared by Pages, Numbers, and Keynote text storage.
#[allow(
    clippy::struct_excessive_bools,
    reason = "Flow mirrors five independent native paragraph inspector controls."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Flow {
    keep_lines_together: bool,
    keep_with_next: bool,
    start_on_new_page: bool,
    prevent_widow_orphan_lines: bool,
    hyphenation: Hyphenation,
}

impl Flow {
    /// Construct the native effective defaults.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            keep_lines_together: false,
            keep_with_next: false,
            start_on_new_page: false,
            prevent_widow_orphan_lines: true,
            hyphenation: Hyphenation::Automatic,
        }
    }

    /// Enable or disable keeping all lines of this paragraph together.
    #[must_use]
    pub const fn with_keep_lines_together(mut self, enabled: bool) -> Self {
        self.keep_lines_together = enabled;
        self
    }

    /// Enable or disable keeping this paragraph with the following one.
    #[must_use]
    pub const fn with_keep_with_next(mut self, enabled: bool) -> Self {
        self.keep_with_next = enabled;
        self
    }

    /// Enable or disable starting this paragraph on a new page.
    #[must_use]
    pub const fn with_start_on_new_page(mut self, enabled: bool) -> Self {
        self.start_on_new_page = enabled;
        self
    }

    /// Enable or disable widow/orphan line prevention.
    #[must_use]
    pub const fn with_prevent_widow_orphan_lines(mut self, enabled: bool) -> Self {
        self.prevent_widow_orphan_lines = enabled;
        self
    }

    /// Select the paragraph's automatic-hyphenation policy.
    #[must_use]
    pub const fn with_hyphenation(mut self, hyphenation: Hyphenation) -> Self {
        self.hyphenation = hyphenation;
        self
    }

    /// Whether all paragraph lines stay together.
    #[must_use]
    pub const fn keeps_lines_together(self) -> bool {
        self.keep_lines_together
    }

    /// Whether this paragraph stays with the following paragraph.
    #[must_use]
    pub const fn keeps_with_next(self) -> bool {
        self.keep_with_next
    }

    /// Whether this paragraph starts on a new page.
    #[must_use]
    pub const fn starts_on_new_page(self) -> bool {
        self.start_on_new_page
    }

    /// Whether widow/orphan line prevention is enabled.
    #[must_use]
    pub const fn prevents_widow_orphan_lines(self) -> bool {
        self.prevent_widow_orphan_lines
    }

    /// Return the automatic-hyphenation policy.
    #[must_use]
    pub const fn hyphenation(self) -> Hyphenation {
        self.hyphenation
    }
}

impl Default for Flow {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_and_builders_match_the_native_inspector() {
        let defaults = Flow::default();
        assert!(!defaults.keeps_lines_together());
        assert!(!defaults.keeps_with_next());
        assert!(!defaults.starts_on_new_page());
        assert!(defaults.prevents_widow_orphan_lines());
        assert_eq!(defaults.hyphenation(), Hyphenation::Automatic);

        let custom = defaults
            .with_keep_lines_together(true)
            .with_keep_with_next(true)
            .with_start_on_new_page(true)
            .with_prevent_widow_orphan_lines(false)
            .with_hyphenation(Hyphenation::Prevented);
        assert!(custom.keeps_lines_together());
        assert!(custom.keeps_with_next());
        assert!(custom.starts_on_new_page());
        assert!(!custom.prevents_widow_orphan_lines());
        assert_eq!(custom.hyphenation(), Hyphenation::Prevented);
    }
}
