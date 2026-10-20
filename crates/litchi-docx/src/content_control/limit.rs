//! Typed budgets of the content-control family (record 0355a).

/// The bounded resources of the content-control family, carried by [`crate::Error::ContentControlLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Aggregate CRC byte.
    AggregateCrcByte,
    /// Content-control attribute count.
    AttributeCount,
    /// Binding count.
    BindingCount,
    /// Content-control count.
    Count,
    /// Event count.
    EventCount,
    /// Input bytes.
    InputBytes,
    /// List-item count.
    ListItemCount,
    /// Content-control MCE depth.
    MceDepth,
    /// Content-control MCE event.
    MceEvent,
    /// MCE metadata bytes.
    MceMetadataBytes,
    /// MCE namespace bindings.
    MceNamespaceBindings,
    /// Content-control MCE offset.
    MceOffset,
    /// MCE output bytes.
    MceOutputBytes,
    /// MCE validation event count.
    MceValidationEventCount,
    /// Metadata bytes.
    MetadataBytes,
    /// Content-control mutation.
    Mutation,
    /// Aggregate content-control output.
    Output,
    /// Per-control list-item count.
    PerControlListItemCount,
    /// Content-control source count.
    SourceCount,
    /// Content-control source event.
    SourceEvent,
    /// Unique CRC part.
    UniqueCrcPart,
    /// Unique Custom XML payload.
    UniqueCustomXmlPayload,
    /// XML depth.
    XmlDepth,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::AggregateCrcByte => "aggregate CRC byte",
            Self::AttributeCount => "content-control attribute count",
            Self::BindingCount => "binding count",
            Self::Count => "content-control count",
            Self::EventCount => "event count",
            Self::InputBytes => "input bytes",
            Self::ListItemCount => "list-item count",
            Self::MceDepth => "content-control MCE depth",
            Self::MceEvent => "content-control MCE event",
            Self::MceMetadataBytes => "MCE metadata bytes",
            Self::MceNamespaceBindings => "MCE namespace bindings",
            Self::MceOffset => "content-control MCE offset",
            Self::MceOutputBytes => "MCE output bytes",
            Self::MceValidationEventCount => "MCE validation event count",
            Self::MetadataBytes => "metadata bytes",
            Self::Mutation => "content-control mutation",
            Self::Output => "aggregate content-control output",
            Self::PerControlListItemCount => "per-control list-item count",
            Self::SourceCount => "content-control source count",
            Self::SourceEvent => "content-control source event",
            Self::UniqueCrcPart => "unique CRC part",
            Self::UniqueCustomXmlPayload => "unique Custom XML payload",
            Self::XmlDepth => "XML depth",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
