//! Typed budgets of the web-settings family (record 0355a).

/// The bounded resources of the web-settings family, carried by [`crate::Error::WebLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Web frameset depth.
    FramesetDepth,
    /// HTML division depth.
    HtmlDivisionDepth,
    /// Web-settings XML depth.
    SettingsXmlDepth,
    /// Web XML depth.
    XmlDepth,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::FramesetDepth => "web frameset depth",
            Self::HtmlDivisionDepth => "HTML division depth",
            Self::SettingsXmlDepth => "web-settings XML depth",
            Self::XmlDepth => "web XML depth",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
