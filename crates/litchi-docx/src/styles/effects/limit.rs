//! Typed budgets of the stylesWithEffects family (record 0355a).

/// The bounded resources of the stylesWithEffects family, carried by [`crate::Error::StylesWithEffectsLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// StylesWithEffects XML bytes.
    XmlBytes,
    /// StylesWithEffects XML depth.
    XmlDepth,
    /// StylesWithEffects XML events.
    XmlEvents,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::XmlBytes => "stylesWithEffects XML bytes",
            Self::XmlDepth => "stylesWithEffects XML depth",
            Self::XmlEvents => "stylesWithEffects XML events",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
