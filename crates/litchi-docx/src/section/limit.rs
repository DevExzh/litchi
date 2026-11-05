//! Typed section budgets (record 0355a).

/// The bounded resources of section inventories and layouts, carried by [`crate::Error::SectionInventoryLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Section bytes.
    Bytes,
    /// Header/footer reference bytes.
    HeaderFooterReferenceBytes,
    /// Input bytes.
    InputBytes,
    /// Paragraphs.
    Paragraphs,
    /// Sections.
    Sections,
    /// XML depth.
    XmlDepth,
    /// XML events.
    XmlEvents,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Bytes => "section bytes",
            Self::HeaderFooterReferenceBytes => "header/footer reference bytes",
            Self::InputBytes => "input bytes",
            Self::Paragraphs => "paragraphs",
            Self::Sections => "sections",
            Self::XmlDepth => "XML depth",
            Self::XmlEvents => "XML events",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
