//! Typed budgets of the settings family (record 0355a).

/// The bounded resources of the settings family, carried by [`crate::Error::SettingsLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Settings XML bytes for document variables.
    DocumentVariableXmlBytes,
    /// Document-variable XML nesting depth.
    DocumentVariableDepth,
    /// Number of document variables.
    DocumentVariableCount,
    /// Unicode scalar values in one document-variable name.
    DocumentVariableNameChars,
    /// Unicode scalar values in one document-variable value.
    DocumentVariableValueChars,
    /// Settings extensions.
    Extensions,
    /// Opaque settings extension bytes.
    OpaqueExtensionBytes,
    /// Opaque settings extension depth.
    OpaqueExtensionDepth,
    /// Opaque settings extension nodes.
    OpaqueExtensionNodes,
    /// Settings XML bytes.
    XmlBytes,
    /// Settings XML depth.
    XmlDepth,
    /// Settings XML nodes.
    XmlNodes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::DocumentVariableXmlBytes => "settings XML bytes for document variables",
            Self::DocumentVariableDepth => "document-variable XML depth",
            Self::DocumentVariableCount => "document-variable count",
            Self::DocumentVariableNameChars => "document-variable name characters",
            Self::DocumentVariableValueChars => "document-variable value characters",
            Self::Extensions => "settings extensions",
            Self::OpaqueExtensionBytes => "opaque settings extension bytes",
            Self::OpaqueExtensionDepth => "opaque settings extension depth",
            Self::OpaqueExtensionNodes => "opaque settings extension nodes",
            Self::XmlBytes => "settings XML bytes",
            Self::XmlDepth => "settings XML depth",
            Self::XmlNodes => "settings XML nodes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
