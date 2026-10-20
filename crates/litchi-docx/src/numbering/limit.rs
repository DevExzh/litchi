//! Typed budgets of the numbering family (record 0355a).

/// The bounded resources of the numbering family, carried by [`crate::Error::NumberingLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Numbering XML bytes.
    XmlBytes,
    /// Numbering XML depth.
    XmlDepth,
    /// Numbering XML elements.
    XmlNodes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::XmlBytes => "numbering XML bytes",
            Self::XmlDepth => "numbering XML depth",
            Self::XmlNodes => "numbering XML elements",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
