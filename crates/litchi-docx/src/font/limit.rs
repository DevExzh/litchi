//! Typed budgets of the font-table family (record 0355a).

/// The bounded resources of the font-table family, carried by [`crate::Error::FontLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Embedded font bytes.
    EmbeddedFontBytes,
    /// Font-table string bytes.
    StringBytes,
    /// Font-table text bytes.
    Text,
    /// Font-usage XML node.
    XmlNode,
    /// Font-table XML structure.
    XmlStructure,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::EmbeddedFontBytes => "embedded font bytes",
            Self::StringBytes => "font-table string bytes",
            Self::Text => "font-table text bytes",
            Self::XmlNode => "font-usage XML node",
            Self::XmlStructure => "font-table XML structure",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
