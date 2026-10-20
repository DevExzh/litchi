//! Typed budgets of the ink family (record 0355a).

/// The bounded resources of the ink family, carried by [`crate::Error::InkLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Ink anchors.
    Anchors,
    /// Annotations.
    Annotations,
    /// Ink authored run bytes.
    AuthoredRunBytes,
    /// Ink authored story changes.
    AuthoredStoryChanges,
    /// Ink authored targets.
    AuthoredTargets,
    /// Ink candidate anchors.
    CandidateAnchors,
    /// Drawing identifiers.
    DrawingIdentifiers,
    /// Durable annotation position.
    DurableAnnotationPosition,
    /// Durable closure bytes.
    DurableClosureBytes,
    /// Durable closure metadata bytes.
    DurableClosureMetadataBytes,
    /// Durable closure records.
    DurableClosureRecords,
    /// Durable closure relationships.
    DurableClosureRelationships,
    /// Durable content type.
    DurableContentType,
    /// Durable fallback image bytes.
    DurableFallbackImageBytes,
    /// Durable geometry height.
    DurableGeometryHeight,
    /// Durable geometry width.
    DurableGeometryWidth,
    /// Durable intent bytes.
    DurableIntentBytes,
    /// Durable intent count.
    DurableIntentCount,
    /// Durable intent records.
    DurableIntentRecords,
    /// Durable member name.
    DurableMemberName,
    /// Durable package XML.
    DurablePackageXml,
    /// Durable paragraph position.
    DurableParagraphPosition,
    /// Durable part payload.
    DurablePartPayload,
    /// Durable prepared Ink bytes.
    DurablePreparedInkBytes,
    /// Durable relationship bytes.
    DurableRelationshipBytes,
    /// Durable restore bytes.
    DurableRestoreBytes,
    /// Durable restore closure bytes.
    DurableRestoreClosureBytes,
    /// Durable restore intent bytes.
    DurableRestoreIntentBytes,
    /// Durable story position.
    DurableStoryPosition,
    /// Edit operations.
    EditOperations,
    /// Edit package bytes.
    EditPackageBytes,
    /// Edit staged bytes.
    EditStagedBytes,
    /// Fallback image bytes.
    FallbackImageBytes,
    /// Ink fallback retarget targets.
    FallbackRetargetTargets,
    /// Ink fallback retargets.
    FallbackRetargets,
    /// Fresh Ink relationship IDs.
    FreshInkRelationshipIDs,
    /// Fresh Ink target names.
    FreshInkTargetNames,
    /// Ink graph package parts.
    GraphPackageParts,
    /// Ink graph publication token bytes.
    GraphPublicationTokenBytes,
    /// Ink graph relationships.
    GraphRelationships,
    /// Ink graph target payload bytes.
    GraphTargetPayloadBytes,
    /// Ink graph topology bytes.
    GraphTopologyBytes,
    /// DOCX ink host candidate anchors.
    HostCandidateAnchors,
    /// Host output bytes.
    HostOutputBytes,
    /// Ink insertion plans.
    InsertionPlans,
    /// MCE choices.
    MceChoices,
    /// MCE fallbacks.
    MceFallbacks,
    /// MCE marked XML bytes.
    MceMarkedXmlBytes,
    /// MCE output bytes.
    MceOutputBytes,
    /// Ink owner relationships.
    OwnerRelationships,
    /// Package parts.
    PackageParts,
    /// DOCX Ink paragraph candidates.
    ParagraphCandidates,
    /// Payload bytes.
    PayloadBytes,
    /// Payload prolog events.
    PayloadPrologEvents,
    /// DOCX Ink profile attribute bytes.
    ProfileAttributeBytes,
    /// DOCX Ink profile attributes.
    ProfileAttributes,
    /// DOCX Ink profile child elements.
    ProfileChildElements,
    /// DOCX Ink profile decoded attribute bytes.
    ProfileDecodedAttributeBytes,
    /// DOCX Ink profile definition records.
    ProfileDefinitionRecords,
    /// DOCX Ink profile source bytes.
    ProfileSourceBytes,
    /// DOCX Ink profile trace references.
    ProfileTraceReferences,
    /// DOCX Ink profile XML depth.
    ProfileXmlDepth,
    /// DOCX Ink profile XML nodes.
    ProfileXmlNodes,
    /// Ink relationship metadata bytes.
    RelationshipMetadataBytes,
    /// DOCX ink relationship reference values.
    RelationshipReferenceValues,
    /// Ink relationship retargets.
    RelationshipRetargets,
    /// Relationships.
    Relationships,
    /// Ink released relationship IDs.
    ReleasedRelationshipIDs,
    /// Remaining incoming relationships.
    RemainingIncomingRelationships,
    /// Ink replacement plans.
    ReplacementPlans,
    /// Ink retarget output bytes.
    RetargetOutputBytes,
    /// Ink retarget replacement bytes.
    RetargetReplacementBytes,
    /// Ink retarget targets.
    RetargetTargets,
    /// Ink retarget XML depth.
    RetargetXmlDepth,
    /// Selected relationships.
    SelectedRelationships,
    /// Ink selected removals.
    SelectedRemovals,
    /// DOCX Ink semantic projection.
    SemanticProjection,
    /// Story XML bytes.
    StoryXmlBytes,
    /// Ink structural edits.
    StructuralEdits,
    /// Target reference bytes.
    TargetReferenceBytes,
    /// Total payload bytes.
    TotalPayloadBytes,
    /// DOCX Ink traceFormat channels.
    TraceFormatChannels,
    /// DOCX Ink traceFormat records.
    TraceFormatRecords,
    /// XML attribute.
    XmlAttribute,
    /// XML attribute QName.
    XmlAttributeQName,
    /// XML attributes.
    XmlAttributes,
    /// XML CDATA.
    XmlCdata,
    /// XML comment.
    XmlComment,
    /// XML declaration.
    XmlDeclaration,
    /// XML depth.
    XmlDepth,
    /// XML DTD.
    XmlDtd,
    /// XML element QName.
    XmlElementQName,
    /// XML end QName.
    XmlEndQName,
    /// XML nodes.
    XmlNodes,
    /// XML processing instruction.
    XmlProcessingInstruction,
    /// XML reference.
    XmlReference,
    /// XML text.
    XmlText,
    /// Ink XML updates.
    XmlUpdates,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Anchors => "ink anchors",
            Self::Annotations => "annotations",
            Self::AuthoredRunBytes => "Ink authored run bytes",
            Self::AuthoredStoryChanges => "Ink authored story changes",
            Self::AuthoredTargets => "Ink authored targets",
            Self::CandidateAnchors => "ink candidate anchors",
            Self::DrawingIdentifiers => "drawing identifiers",
            Self::DurableAnnotationPosition => "durable annotation position",
            Self::DurableClosureBytes => "durable closure bytes",
            Self::DurableClosureMetadataBytes => "durable closure metadata bytes",
            Self::DurableClosureRecords => "durable closure records",
            Self::DurableClosureRelationships => "durable closure relationships",
            Self::DurableContentType => "durable content type",
            Self::DurableFallbackImageBytes => "durable fallback image bytes",
            Self::DurableGeometryHeight => "durable geometry height",
            Self::DurableGeometryWidth => "durable geometry width",
            Self::DurableIntentBytes => "durable intent bytes",
            Self::DurableIntentCount => "durable intent count",
            Self::DurableIntentRecords => "durable intent records",
            Self::DurableMemberName => "durable member name",
            Self::DurablePackageXml => "durable package XML",
            Self::DurableParagraphPosition => "durable paragraph position",
            Self::DurablePartPayload => "durable part payload",
            Self::DurablePreparedInkBytes => "durable prepared Ink bytes",
            Self::DurableRelationshipBytes => "durable relationship bytes",
            Self::DurableRestoreBytes => "durable restore bytes",
            Self::DurableRestoreClosureBytes => "durable restore closure bytes",
            Self::DurableRestoreIntentBytes => "durable restore intent bytes",
            Self::DurableStoryPosition => "durable story position",
            Self::EditOperations => "edit operations",
            Self::EditPackageBytes => "edit package bytes",
            Self::EditStagedBytes => "edit staged bytes",
            Self::FallbackImageBytes => "fallback image bytes",
            Self::FallbackRetargetTargets => "Ink fallback retarget targets",
            Self::FallbackRetargets => "Ink fallback retargets",
            Self::FreshInkRelationshipIDs => "fresh Ink relationship IDs",
            Self::FreshInkTargetNames => "fresh Ink target names",
            Self::GraphPackageParts => "Ink graph package parts",
            Self::GraphPublicationTokenBytes => "Ink graph publication token bytes",
            Self::GraphRelationships => "Ink graph relationships",
            Self::GraphTargetPayloadBytes => "Ink graph target payload bytes",
            Self::GraphTopologyBytes => "Ink graph topology bytes",
            Self::HostCandidateAnchors => "DOCX ink host candidate anchors",
            Self::HostOutputBytes => "host output bytes",
            Self::InsertionPlans => "Ink insertion plans",
            Self::MceChoices => "MCE choices",
            Self::MceFallbacks => "MCE fallbacks",
            Self::MceMarkedXmlBytes => "MCE marked XML bytes",
            Self::MceOutputBytes => "MCE output bytes",
            Self::OwnerRelationships => "Ink owner relationships",
            Self::PackageParts => "package parts",
            Self::ParagraphCandidates => "DOCX Ink paragraph candidates",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadPrologEvents => "payload prolog events",
            Self::ProfileAttributeBytes => "DOCX Ink profile attribute bytes",
            Self::ProfileAttributes => "DOCX Ink profile attributes",
            Self::ProfileChildElements => "DOCX Ink profile child elements",
            Self::ProfileDecodedAttributeBytes => "DOCX Ink profile decoded attribute bytes",
            Self::ProfileDefinitionRecords => "DOCX Ink profile definition records",
            Self::ProfileSourceBytes => "DOCX Ink profile source bytes",
            Self::ProfileTraceReferences => "DOCX Ink profile trace references",
            Self::ProfileXmlDepth => "DOCX Ink profile XML depth",
            Self::ProfileXmlNodes => "DOCX Ink profile XML nodes",
            Self::RelationshipMetadataBytes => "Ink relationship metadata bytes",
            Self::RelationshipReferenceValues => "DOCX ink relationship reference values",
            Self::RelationshipRetargets => "Ink relationship retargets",
            Self::Relationships => "relationships",
            Self::ReleasedRelationshipIDs => "Ink released relationship IDs",
            Self::RemainingIncomingRelationships => "remaining incoming relationships",
            Self::ReplacementPlans => "Ink replacement plans",
            Self::RetargetOutputBytes => "Ink retarget output bytes",
            Self::RetargetReplacementBytes => "Ink retarget replacement bytes",
            Self::RetargetTargets => "Ink retarget targets",
            Self::RetargetXmlDepth => "Ink retarget XML depth",
            Self::SelectedRelationships => "selected relationships",
            Self::SelectedRemovals => "Ink selected removals",
            Self::SemanticProjection => "DOCX Ink semantic projection",
            Self::StoryXmlBytes => "story XML bytes",
            Self::StructuralEdits => "Ink structural edits",
            Self::TargetReferenceBytes => "target reference bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::TraceFormatChannels => "DOCX Ink traceFormat channels",
            Self::TraceFormatRecords => "DOCX Ink traceFormat records",
            Self::XmlAttribute => "XML attribute",
            Self::XmlAttributeQName => "XML attribute QName",
            Self::XmlAttributes => "XML attributes",
            Self::XmlCdata => "XML CDATA",
            Self::XmlComment => "XML comment",
            Self::XmlDeclaration => "XML declaration",
            Self::XmlDepth => "XML depth",
            Self::XmlDtd => "XML DTD",
            Self::XmlElementQName => "XML element QName",
            Self::XmlEndQName => "XML end QName",
            Self::XmlNodes => "XML nodes",
            Self::XmlProcessingInstruction => "XML processing instruction",
            Self::XmlReference => "XML reference",
            Self::XmlText => "XML text",
            Self::XmlUpdates => "Ink XML updates",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
