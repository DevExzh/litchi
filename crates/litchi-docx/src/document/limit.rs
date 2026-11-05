//! Typed resource dimensions for main-document transaction failures.

/// The bounded resource reported by a primary document transaction.
///
/// Variants distinguish XML, staging, managed accounting and durable/history
/// dimensions. They carry structural metadata only, never document content.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Conservative staging bound for generated formatting XML and property plans.
    FormattingStagingBytes,
    /// Aggregate payload bytes in a durable transfer graph.
    TransferGraphBytes,
    /// XML bytes.
    XmlBytes,
    /// XML depth.
    XmlDepth,
    /// XML elements.
    XmlElements,
    /// Attributes on one XML element.
    XmlAttributes,
    /// Raw attribute bytes on one XML element.
    XmlAttributeBytes,
    /// Block content control path.
    BlockContentControlPath,
    /// Composite hyperlink replacements.
    CompositeHyperlinkReplacements,
    /// Composite paragraph replacements.
    CompositeParagraphReplacements,
    /// Content control path.
    ContentControlPath,
    /// Document namespace binding count.
    DocumentNamespaceBindingCount,
    /// Document namespace lookup work.
    DocumentNamespaceLookupWork,
    /// Durable nested table path.
    DurableNestedTablePath,
    /// Durable operations.
    DurableOperations,
    /// Durable selector path.
    DurableSelectorPath,
    /// Field bytes.
    FieldBytes,
    /// History dependency bytes.
    HistoryDependencyBytes,
    /// History transition bytes.
    HistoryTransitionBytes,
    /// Inverse operation string bytes.
    InverseOperationStringBytes,
    /// Inverse operation string metadata bytes.
    InverseOperationStringMetadataBytes,
    /// Main relationships.
    MainRelationships,
    /// Managed batch input admission capacity.
    ManagedBatchInputAdmissionCapacity,
    /// Managed batch input admission metadata.
    ManagedBatchInputAdmissionMetadata,
    /// Managed cloned operation string bytes.
    ManagedClonedOperationStringBytes,
    /// Managed cloned operation string metadata bytes.
    ManagedClonedOperationStringMetadataBytes,
    /// Managed namespace lookup work.
    ManagedNamespaceLookupWork,
    /// Managed operation metadata bytes.
    ManagedOperationMetadataBytes,
    /// Managed operation string bytes.
    ManagedOperationStringBytes,
    /// Managed paragraph batch decision map metadata.
    ManagedParagraphBatchDecisionMapMetadata,
    /// Managed paragraph batch decision metadata.
    ManagedParagraphBatchDecisionMetadata,
    /// Managed paragraph fragment bytes.
    ManagedParagraphFragmentBytes,
    /// Managed paragraph fragment count.
    ManagedParagraphFragmentCount,
    /// Managed paragraph fragment memory.
    ManagedParagraphFragmentMemory,
    /// Managed paragraph fragment metadata.
    ManagedParagraphFragmentMetadata,
    /// Managed paragraph operation count.
    ManagedParagraphOperationCount,
    /// Managed paragraph reconstruction plan memory.
    ManagedParagraphReconstructionPlanMemory,
    /// Managed paragraph reconstruction work.
    ManagedParagraphReconstructionWork,
    /// Managed reservation amount.
    ManagedReservationAmount,
    /// Managed scan memory.
    ManagedScanMemory,
    /// Managed scan objects.
    ManagedScanObjects,
    /// Managed scan parser memory.
    ManagedScanParserMemory,
    /// Managed scan work.
    ManagedScanWork,
    /// Managed snapshot index bytes.
    ManagedSnapshotIndexBytes,
    /// Managed snapshot memory.
    ManagedSnapshotMemory,
    /// Managed snapshot metadata bytes.
    ManagedSnapshotMetadataBytes,
    /// Managed snapshot objects.
    ManagedSnapshotObjects,
    /// Managed text-slot count.
    ManagedTextSlotCount,
    /// Managed text-slot memory.
    ManagedTextSlotMemory,
    /// Managed work amount.
    ManagedWorkAmount,
    /// Nested table path.
    NestedTablePath,
    /// Operation string metadata bytes.
    OperationStringMetadataBytes,
    /// Operations.
    Operations,
    /// Paragraph replacement XML bytes.
    ParagraphReplacementXmlBytes,
    /// Paragraphs.
    Paragraphs,
    /// Part relationships.
    PartRelationships,
    /// Parts.
    Parts,
    /// Projected XML bytes.
    ProjectedXmlBytes,
    /// Replacement XML bytes.
    ReplacementXmlBytes,
    /// Replacement text bytes.
    ReplacementTextBytes,
    /// Replacement text characters, conservatively capped by the UTF-8 byte ceiling.
    ReplacementTextCharacters,
    /// Restore artifact bytes.
    RestoreArtifactBytes,
    /// Story hyperlink replacements.
    StoryHyperlinkReplacements,
    /// Rows of a new table.
    TableRows,
    /// Columns of a new table or cells of a new row.
    TableColumns,
    /// Transfer dependency bytes.
    TransferDependencyBytes,
    /// Transfer dependency field bytes.
    TransferDependencyFieldBytes,
    /// Transfer dependency parts.
    TransferDependencyParts,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(match self {
            Self::FormattingStagingBytes => "formatting staging bytes",
            Self::TransferGraphBytes => "transfer graph bytes",
            Self::XmlBytes => "XML bytes",
            Self::XmlDepth => "XML depth",
            Self::XmlElements => "XML elements",
            Self::XmlAttributes => "XML attributes",
            Self::XmlAttributeBytes => "XML attribute bytes",
            Self::BlockContentControlPath => "block content control path",
            Self::CompositeHyperlinkReplacements => "composite hyperlink replacements",
            Self::CompositeParagraphReplacements => "composite paragraph replacements",
            Self::ContentControlPath => "content control path",
            Self::DocumentNamespaceBindingCount => "document namespace binding count",
            Self::DocumentNamespaceLookupWork => "document namespace lookup work",
            Self::DurableNestedTablePath => "durable nested table path",
            Self::DurableOperations => "durable operations",
            Self::DurableSelectorPath => "durable selector path",
            Self::FieldBytes => "field bytes",
            Self::HistoryDependencyBytes => "history dependency bytes",
            Self::HistoryTransitionBytes => "history transition bytes",
            Self::InverseOperationStringBytes => "inverse operation string bytes",
            Self::InverseOperationStringMetadataBytes => "inverse operation string metadata bytes",
            Self::MainRelationships => "main relationships",
            Self::ManagedBatchInputAdmissionCapacity => "managed batch input admission capacity",
            Self::ManagedBatchInputAdmissionMetadata => "managed batch input admission metadata",
            Self::ManagedClonedOperationStringBytes => "managed cloned operation string bytes",
            Self::ManagedClonedOperationStringMetadataBytes => {
                "managed cloned operation string metadata bytes"
            },
            Self::ManagedNamespaceLookupWork => "managed namespace lookup work",
            Self::ManagedOperationMetadataBytes => "managed operation metadata bytes",
            Self::ManagedOperationStringBytes => "managed operation string bytes",
            Self::ManagedParagraphBatchDecisionMapMetadata => {
                "managed paragraph batch decision map metadata"
            },
            Self::ManagedParagraphBatchDecisionMetadata => {
                "managed paragraph batch decision metadata"
            },
            Self::ManagedParagraphFragmentBytes => "managed paragraph fragment bytes",
            Self::ManagedParagraphFragmentCount => "managed paragraph fragment count",
            Self::ManagedParagraphFragmentMemory => "managed paragraph fragment memory",
            Self::ManagedParagraphFragmentMetadata => "managed paragraph fragment metadata",
            Self::ManagedParagraphOperationCount => "managed paragraph operation count",
            Self::ManagedParagraphReconstructionPlanMemory => {
                "managed paragraph reconstruction plan memory"
            },
            Self::ManagedParagraphReconstructionWork => "managed paragraph reconstruction work",
            Self::ManagedReservationAmount => "managed reservation amount",
            Self::ManagedScanMemory => "managed scan memory",
            Self::ManagedScanObjects => "managed scan objects",
            Self::ManagedScanParserMemory => "managed scan parser memory",
            Self::ManagedScanWork => "managed scan work",
            Self::ManagedSnapshotIndexBytes => "managed snapshot index bytes",
            Self::ManagedSnapshotMemory => "managed snapshot memory",
            Self::ManagedSnapshotMetadataBytes => "managed snapshot metadata bytes",
            Self::ManagedSnapshotObjects => "managed snapshot objects",
            Self::ManagedTextSlotCount => "managed text-slot count",
            Self::ManagedTextSlotMemory => "managed text-slot memory",
            Self::ManagedWorkAmount => "managed work amount",
            Self::NestedTablePath => "nested table path",
            Self::OperationStringMetadataBytes => "operation string metadata bytes",
            Self::Operations => "operations",
            Self::ParagraphReplacementXmlBytes => "paragraph replacement XML bytes",
            Self::Paragraphs => "paragraphs",
            Self::PartRelationships => "part relationships",
            Self::Parts => "parts",
            Self::ProjectedXmlBytes => "projected XML bytes",
            Self::ReplacementXmlBytes => "replacement XML bytes",
            Self::ReplacementTextBytes => "replacement text bytes",
            Self::ReplacementTextCharacters => "replacement text characters",
            Self::RestoreArtifactBytes => "restore artifact bytes",
            Self::StoryHyperlinkReplacements => "story hyperlink replacements",
            Self::TableRows => "table rows",
            Self::TableColumns => "table columns",
            Self::TransferDependencyBytes => "transfer dependency bytes",
            Self::TransferDependencyFieldBytes => "transfer dependency field bytes",
            Self::TransferDependencyParts => "transfer dependency parts",
        })
    }
}
