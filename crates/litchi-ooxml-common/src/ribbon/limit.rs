//! The typed budgets of the Ribbon (customUI) family (record 0360).

/// A Ribbon (customUI) family budget that was exceeded.
///
/// Every limit error of the [`crate::ribbon`] module is
/// [`crate::Error::Limit`] with one of these kinds; its display is the
/// resource's human-readable name.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// The bound on Ribbon image garbage-collection edges.
    ImageGarbageCollectionEdges,
    /// The bound on Ribbon image relationships.
    ImageRelationships,
    /// The bound on Ribbon package graph relationships.
    PackageGraphRelationships,
    /// The bound on Ribbon part-name allocation bytes.
    PartNameAllocationBytes,
    /// The bound on Ribbon part-name allocation scan.
    PartNameAllocationScan,
    /// The bound on Ribbon XML bytes.
    XmlBytes,
    /// The bound on Ribbon XML depth.
    XmlDepth,
    /// The bound on Ribbon XML nodes.
    XmlNodes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::ImageGarbageCollectionEdges => "Ribbon image garbage-collection edges",
            Self::ImageRelationships => "Ribbon image relationships",
            Self::PackageGraphRelationships => "Ribbon package graph relationships",
            Self::PartNameAllocationBytes => "Ribbon part-name allocation bytes",
            Self::PartNameAllocationScan => "Ribbon part-name allocation scan",
            Self::XmlBytes => "Ribbon XML bytes",
            Self::XmlDepth => "Ribbon XML depth",
            Self::XmlNodes => "Ribbon XML nodes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
