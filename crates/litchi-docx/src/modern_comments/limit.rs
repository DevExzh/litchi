//! Typed budgets of the modern Word comments raw codec.

/// A bounded resource used by a modern-comment parser or writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Serialized modern-comment part bytes.
    XmlBytes,
    /// Serialized writer output bytes.
    OutputBytes,
    /// XML nesting depth.
    Depth,
    /// Parsed XML nodes or metadata items.
    Items,
    /// Decoded XML string bytes.
    StringBytes,
}

impl LimitKind {
    /// Stable human-readable resource name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::XmlBytes => "modern-comment XML bytes",
            Self::OutputBytes => "modern-comment output bytes",
            Self::Depth => "modern-comment XML depth",
            Self::Items => "modern-comment items",
            Self::StringBytes => "modern-comment string bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
