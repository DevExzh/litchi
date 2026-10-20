//! Typed budgets of the mail-merge family (record 0355a).

/// The bounded resources of the mail-merge family, carried by [`crate::Error::MailMergeLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Odso field maps.
    FieldMaps,
    /// Mail-merge recipients.
    Recipients,
    /// Mail-merge XML attributes on one node.
    XmlAttributes,
    /// Mail-merge XML nodes.
    XmlNodes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::FieldMaps => "odso field maps",
            Self::Recipients => "mail-merge recipients",
            Self::XmlAttributes => "mail-merge XML attributes on one node",
            Self::XmlNodes => "mail-merge XML nodes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
