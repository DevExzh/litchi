//! The typed budgets of the Custom XML family (record 0360).

/// A Custom XML budget that a read, an edit or a patch exceeded.
///
/// Every limit error of this module is
/// [`crate::Error::CustomXmlLimit`] with one of these kinds; its display
/// is the resource's human-readable name.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Custom XML element nesting depth.
    Depth,
    /// Custom XML elements of one payload.
    Elements,
    /// Custom XML data items of one package.
    Items,
    /// Bytes of one Custom XML data part.
    PayloadBytes,
    /// Bytes of one Custom XML properties part.
    PropertiesBytes,
    /// Bytes of one string inside a properties part.
    PropertyStringBytes,
    /// Schema references of one properties part.
    SchemaReferences,
    /// Bytes of a serialized properties part.
    SerializedPropertiesBytes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Depth => "custom XML depth",
            Self::Elements => "custom XML elements",
            Self::Items => "custom XML items",
            Self::PayloadBytes => "custom XML payload bytes",
            Self::PropertiesBytes => "custom XML properties bytes",
            Self::PropertyStringBytes => "custom XML property string bytes",
            Self::SchemaReferences => "custom XML schema references",
            Self::SerializedPropertiesBytes => "custom XML serialized properties bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
