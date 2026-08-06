//! Closed semantic values for the Word 2012 `collapsed` element.

/// The value of a Word 2012 `w12:collapsed` element.
///
/// Absence of the element is represented by `Option<Collapsed>` at the
/// paragraph boundary. Keeping the wire value separate from absence prevents
/// an explicit false value from being confused with an omitted extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Collapsed {
    /// Subsequent paragraphs with a deeper heading level are displayed
    /// collapsed.
    Enabled,
    /// The paragraph explicitly opts out of collapsing subsequent paragraphs.
    Disabled,
}

impl Collapsed {
    /// Construct the enabled state.
    #[must_use]
    pub const fn enabled() -> Self {
        Self::Enabled
    }

    /// Construct the disabled state.
    #[must_use]
    pub const fn disabled() -> Self {
        Self::Disabled
    }

    /// Return the semantic Boolean value.
    #[must_use]
    pub const fn is_enabled(self) -> bool {
        matches!(self, Self::Enabled)
    }

    /// Return the canonical `w12:val` lexical value.
    #[must_use]
    pub(crate) const fn as_xml(self) -> &'static str {
        if self.is_enabled() { "true" } else { "false" }
    }
}

impl From<bool> for Collapsed {
    fn from(value: bool) -> Self {
        if value { Self::Enabled } else { Self::Disabled }
    }
}

impl From<Collapsed> for bool {
    fn from(value: Collapsed) -> Self {
        value.is_enabled()
    }
}
