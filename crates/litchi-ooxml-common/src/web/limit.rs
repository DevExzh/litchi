//! The typed budgets of the Web Extensions (task-pane) family (record 0360).

/// A Web Extensions (task-pane) family budget that was exceeded.
///
/// Every limit error of the [`crate::web`] module is
/// [`crate::Error::Limit`] with one of these kinds; its display is the
/// resource's human-readable name.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// The bound on Web Extensions add-in XML.
    AddInXml,
    /// The bound on aggregate web extension image bytes.
    AggregateImageBytes,
    /// The bound on aggregate web extension snapshot bytes.
    AggregateSnapshotBytes,
    /// The bound on aggregate web extension XML bytes.
    AggregateXmlBytes,
    /// The bound on alternate reference.
    AlternateReferenceItems,
    /// The bound on alternate references.
    AlternateReferences,
    /// The bound on authored web extension package metadata bytes.
    AuthoredPackageMetadataBytes,
    /// The bound on authored web extension XML bytes.
    AuthoredXmlBytes,
    /// The bound on background runtime ID.
    BackgroundRuntimeId,
    /// The bound on binding.
    BindingItems,
    /// The bound on web extension bindings.
    Bindings,
    /// The bound on content-types XML.
    ContentTypesXml,
    /// The bound on custom-function ID.
    CustomFunctionId,
    /// The bound on custom function ID bytes.
    CustomFunctionIdBytes,
    /// The bound on custom-function ID count.
    CustomFunctionIdCount,
    /// The bound on custom-function IDs.
    CustomFunctionIds,
    /// The bound on Web Extensions custom-function index.
    CustomFunctionIndex,
    /// The bound on custom function IDs.
    CustomFunctionListIds,
    /// The bound on custom-function OfficeArt extensions.
    CustomFunctionOfficeArtExtensions,
    /// The bound on custom-function payloads.
    CustomFunctionPayloads,
    /// The bound on custom-function XML.
    CustomFunctionXml,
    /// The bound on custom-function XML bytes.
    CustomFunctionXmlBytes,
    /// The bound on custom-function XML depth.
    CustomFunctionXmlDepth,
    /// The bound on custom-function XML nodes.
    CustomFunctionXmlNodes,
    /// The bound on web extension decoded string bytes.
    DecodedStringBytes,
    /// The bound on Web Extensions dock state.
    DockState,
    /// The bound on Web Extensions durable authored panes.
    DurableAuthoredPanes,
    /// The bound on Web Extensions durable closure.
    DurableClosure,
    /// The bound on Web Extensions durable closure bytes.
    DurableClosureBytes,
    /// The bound on Web Extensions durable closure member name.
    DurableClosureMemberName,
    /// The bound on Web Extensions durable closure payload.
    DurableClosurePayload,
    /// The bound on Web Extensions durable closure records.
    DurableClosureRecords,
    /// The bound on Web Extensions durable content type.
    DurableContentType,
    /// The bound on Web Extensions durable decoded strings.
    DurableDecodedStrings,
    /// The bound on Web Extensions durable image index.
    DurableImageIndex,
    /// The bound on Web Extensions durable image ledger.
    DurableImageLedger,
    /// The bound on Web Extensions durable image payload bytes.
    DurableImagePayloadBytes,
    /// The bound on Web Extensions durable intent.
    DurableIntent,
    /// The bound on Web Extensions durable intent bytes.
    DurableIntentBytes,
    /// The bound on Web Extensions durable operation count.
    DurableOperationCount,
    /// The bound on Web Extensions durable operations.
    DurableOperations,
    /// The bound on Web Extensions durable pane count.
    DurablePaneCount,
    /// The bound on Web Extensions durable panes.
    DurablePanes,
    /// The bound on Web Extensions durable proof bytes.
    DurableProofBytes,
    /// The bound on Web Extensions durable proof parts.
    DurableProofParts,
    /// The bound on Web Extensions durable relationship XML.
    DurableRelationshipXml,
    /// The bound on Web Extensions durable restore bytes.
    DurableRestoreBytes,
    /// The bound on Web Extensions durable snapshot resources.
    DurableSnapshotResources,
    /// The bound on Web Extensions durable snapshot staging.
    DurableSnapshotStaging,
    /// The bound on Web Extensions durable XML payload bytes.
    DurableXmlPayloadBytes,
    /// The bound on web extension extLst XML bytes.
    ExtLstXmlBytes,
    /// The bound on Web Extensions external snapshot target.
    ExternalSnapshotTarget,
    /// The bound on indexed web extension package metadata bytes.
    IndexedPackageMetadataBytes,
    /// The bound on OfficeArt extension.
    OfficeArtExtensionItems,
    /// The bound on Web Extensions owner ID.
    OwnerId,
    /// The bound on Web Extensions custom-function ID.
    PackageCustomFunctionId,
    /// The bound on Web Extensions custom-function IDs.
    PackageCustomFunctionIds,
    /// The bound on package parts.
    PackageParts,
    /// The bound on package relationships.
    PackageRelationships,
    /// The bound on Web Extensions snapshot bytes.
    PackageSnapshotBytes,
    /// The bound on Web Extensions XML bytes.
    PackageXmlBytes,
    /// The bound on Web Extensions pane extension-list XML.
    PaneExtensionListXml,
    /// The bound on Web Extensions pane index.
    PaneIndex,
    /// The bound on Web Extensions pane relationship ID.
    PaneRelationshipId,
    /// The bound on web extension pane relationship IDs.
    PaneRelationshipIds,
    /// The bound on web extension panes.
    Panes,
    /// The bound on web extension part allocations.
    PartAllocations,
    /// The bound on web extension part deletions.
    PartDeletions,
    /// The bound on web extension part-name allocation probes.
    PartNameAllocationProbes,
    /// The bound on Web Extensions patch lexical scope.
    PatchLexicalScope,
    /// The bound on Web Extensions patch lexical XML.
    PatchLexicalXml,
    /// The bound on Web Extensions patch parts.
    PatchParts,
    /// The bound on Web Extensions patch relationship tokens.
    PatchRelationshipTokens,
    /// The bound on web extension properties.
    Properties,
    /// The bound on property.
    PropertyItems,
    /// The bound on web extension relationship IDs.
    RelationshipIds,
    /// The bound on retained web extension fragment bytes.
    RetainedFragmentBytes,
    /// The bound on retained web extension namespace bytes.
    RetainedNamespaceBytes,
    /// The bound on retained web extension namespace entries.
    RetainedNamespaceEntries,
    /// The bound on retained web extension string bytes.
    RetainedStringBytes,
    /// The bound on web extension snapshot bytes.
    SnapshotBytes,
    /// The bound on Web Extensions snapshot content type.
    SnapshotContentType,
    /// The bound on snapshot effect.
    SnapshotEffectItems,
    /// The bound on snapshot effects.
    SnapshotEffects,
    /// The bound on Web Extensions snapshot part name.
    SnapshotPartName,
    /// The bound on Web Extensions snapshot relationship ID.
    SnapshotRelationshipId,
    /// The bound on snapshot relationship IDs.
    SnapshotRelationshipIds,
    /// The bound on Web Extensions snapshot resource count.
    SnapshotResourceCount,
    /// The bound on Web Extensions source attribute edits.
    SourceAttributeEdits,
    /// The bound on Web Extensions source tag attributes.
    SourceTagAttributes,
    /// The bound on task pane.
    TaskPaneItems,
    /// The bound on Web Extensions task-pane XML.
    TaskPaneXml,
    /// The bound on task-pane XML bytes.
    TaskPaneXmlBytes,
    /// The bound on web extension XML bytes.
    XmlBytes,
    /// The bound on web extension XML depth.
    XmlDepth,
    /// The bound on web extension XML namespace bindings.
    XmlNamespaceBindings,
    /// The bound on web extension XML nodes.
    XmlNodes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::AddInXml => "Web Extensions add-in XML",
            Self::AggregateImageBytes => "aggregate web extension image bytes",
            Self::AggregateSnapshotBytes => "aggregate web extension snapshot bytes",
            Self::AggregateXmlBytes => "aggregate web extension XML bytes",
            Self::AlternateReferenceItems => "alternate reference",
            Self::AlternateReferences => "alternate references",
            Self::AuthoredPackageMetadataBytes => "authored web extension package metadata bytes",
            Self::AuthoredXmlBytes => "authored web extension XML bytes",
            Self::BackgroundRuntimeId => "background runtime ID",
            Self::BindingItems => "binding",
            Self::Bindings => "web extension bindings",
            Self::ContentTypesXml => "content-types XML",
            Self::CustomFunctionId => "custom-function ID",
            Self::CustomFunctionIdBytes => "custom function ID bytes",
            Self::CustomFunctionIdCount => "custom-function ID count",
            Self::CustomFunctionIds => "custom-function IDs",
            Self::CustomFunctionIndex => "Web Extensions custom-function index",
            Self::CustomFunctionListIds => "custom function IDs",
            Self::CustomFunctionOfficeArtExtensions => "custom-function OfficeArt extensions",
            Self::CustomFunctionPayloads => "custom-function payloads",
            Self::CustomFunctionXml => "custom-function XML",
            Self::CustomFunctionXmlBytes => "custom-function XML bytes",
            Self::CustomFunctionXmlDepth => "custom-function XML depth",
            Self::CustomFunctionXmlNodes => "custom-function XML nodes",
            Self::DecodedStringBytes => "web extension decoded string bytes",
            Self::DockState => "Web Extensions dock state",
            Self::DurableAuthoredPanes => "Web Extensions durable authored panes",
            Self::DurableClosure => "Web Extensions durable closure",
            Self::DurableClosureBytes => "Web Extensions durable closure bytes",
            Self::DurableClosureMemberName => "Web Extensions durable closure member name",
            Self::DurableClosurePayload => "Web Extensions durable closure payload",
            Self::DurableClosureRecords => "Web Extensions durable closure records",
            Self::DurableContentType => "Web Extensions durable content type",
            Self::DurableDecodedStrings => "Web Extensions durable decoded strings",
            Self::DurableImageIndex => "Web Extensions durable image index",
            Self::DurableImageLedger => "Web Extensions durable image ledger",
            Self::DurableImagePayloadBytes => "Web Extensions durable image payload bytes",
            Self::DurableIntent => "Web Extensions durable intent",
            Self::DurableIntentBytes => "Web Extensions durable intent bytes",
            Self::DurableOperationCount => "Web Extensions durable operation count",
            Self::DurableOperations => "Web Extensions durable operations",
            Self::DurablePaneCount => "Web Extensions durable pane count",
            Self::DurablePanes => "Web Extensions durable panes",
            Self::DurableProofBytes => "Web Extensions durable proof bytes",
            Self::DurableProofParts => "Web Extensions durable proof parts",
            Self::DurableRelationshipXml => "Web Extensions durable relationship XML",
            Self::DurableRestoreBytes => "Web Extensions durable restore bytes",
            Self::DurableSnapshotResources => "Web Extensions durable snapshot resources",
            Self::DurableSnapshotStaging => "Web Extensions durable snapshot staging",
            Self::DurableXmlPayloadBytes => "Web Extensions durable XML payload bytes",
            Self::ExtLstXmlBytes => "web extension extLst XML bytes",
            Self::ExternalSnapshotTarget => "Web Extensions external snapshot target",
            Self::IndexedPackageMetadataBytes => "indexed web extension package metadata bytes",
            Self::OfficeArtExtensionItems => "OfficeArt extension",
            Self::OwnerId => "Web Extensions owner ID",
            Self::PackageCustomFunctionId => "Web Extensions custom-function ID",
            Self::PackageCustomFunctionIds => "Web Extensions custom-function IDs",
            Self::PackageParts => "package parts",
            Self::PackageRelationships => "package relationships",
            Self::PackageSnapshotBytes => "Web Extensions snapshot bytes",
            Self::PackageXmlBytes => "Web Extensions XML bytes",
            Self::PaneExtensionListXml => "Web Extensions pane extension-list XML",
            Self::PaneIndex => "Web Extensions pane index",
            Self::PaneRelationshipId => "Web Extensions pane relationship ID",
            Self::PaneRelationshipIds => "web extension pane relationship IDs",
            Self::Panes => "web extension panes",
            Self::PartAllocations => "web extension part allocations",
            Self::PartDeletions => "web extension part deletions",
            Self::PartNameAllocationProbes => "web extension part-name allocation probes",
            Self::PatchLexicalScope => "Web Extensions patch lexical scope",
            Self::PatchLexicalXml => "Web Extensions patch lexical XML",
            Self::PatchParts => "Web Extensions patch parts",
            Self::PatchRelationshipTokens => "Web Extensions patch relationship tokens",
            Self::Properties => "web extension properties",
            Self::PropertyItems => "property",
            Self::RelationshipIds => "web extension relationship IDs",
            Self::RetainedFragmentBytes => "retained web extension fragment bytes",
            Self::RetainedNamespaceBytes => "retained web extension namespace bytes",
            Self::RetainedNamespaceEntries => "retained web extension namespace entries",
            Self::RetainedStringBytes => "retained web extension string bytes",
            Self::SnapshotBytes => "web extension snapshot bytes",
            Self::SnapshotContentType => "Web Extensions snapshot content type",
            Self::SnapshotEffectItems => "snapshot effect",
            Self::SnapshotEffects => "snapshot effects",
            Self::SnapshotPartName => "Web Extensions snapshot part name",
            Self::SnapshotRelationshipId => "Web Extensions snapshot relationship ID",
            Self::SnapshotRelationshipIds => "snapshot relationship IDs",
            Self::SnapshotResourceCount => "Web Extensions snapshot resource count",
            Self::SourceAttributeEdits => "Web Extensions source attribute edits",
            Self::SourceTagAttributes => "Web Extensions source tag attributes",
            Self::TaskPaneItems => "task pane",
            Self::TaskPaneXml => "Web Extensions task-pane XML",
            Self::TaskPaneXmlBytes => "task-pane XML bytes",
            Self::XmlBytes => "web extension XML bytes",
            Self::XmlDepth => "web extension XML depth",
            Self::XmlNamespaceBindings => "web extension XML namespace bindings",
            Self::XmlNodes => "web extension XML nodes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
