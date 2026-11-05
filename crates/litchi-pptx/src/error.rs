//! Typed PPTX failures.

use std::collections::TryReserveError;

use thiserror::Error;

/// Result of a PPTX operation.
pub type Result<T> = std::result::Result<T, Error>;

/// A bounded resource dimension used by a PresentationML error.
///
/// The primary opened-presentation edit dimensions remain direct variants so
/// callers can select them through [`crate::opened::Limits`].  Codec and
/// shared-OOXML dimensions are nested to keep their ownership explicit: a
/// codec limit is not an edit limit, and a shared service keeps its own
/// vocabulary.  No untyped resource string crosses the public error boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Changed parts in one primary patch.
    Parts,
    /// Encoded durable-patch bytes.
    PatchBytes,
    /// Replacement text bytes.
    TextBytes,
    /// Retained history entries.
    HistoryEntries,
    /// Aggregate encoded history bytes.
    HistoryBytes,
    /// Retained candidate bytes.
    RetainedCandidateBytes,
    /// Retained markup-compatibility bytes.
    RetainedMceBytes,
    /// A bounded PresentationML codec dimension.
    Codec(CodecLimitKind),
    /// A bounded host-neutral OOXML service dimension.
    Common(litchi_ooxml_common::LimitKind),
}

/// Semantic resource dimensions used by the PPTX codecs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CodecLimitKind {
    /// Bytes admitted from an input buffer.
    InputBytes,
    /// Bytes produced by a codec or writer.
    OutputBytes,
    /// PresentationML or DrawingML XML bytes.
    XmlBytes,
    /// XML nesting depth.
    XmlDepth,
    /// XML elements, nodes, or frames.
    XmlNodes,
    /// XML parser events.
    XmlEvents,
    /// XML attributes.
    XmlAttributes,
    /// XML attribute bytes.
    XmlAttributeBytes,
    /// Decoded string bytes.
    StringBytes,
    /// Text bytes.
    TextBytes,
    /// OPC part bytes.
    PartBytes,
    /// OPC part count.
    Parts,
    /// Relationship count or relationship graph edges.
    Relationships,
    /// Relationship identifier or target bytes.
    RelationshipBytes,
    /// XML entity-reference bytes.
    ReferenceBytes,
    /// Names or name-bearing selectors.
    Names,
    /// Name bytes.
    NameBytes,
    /// Slides or slide references.
    Slides,
    /// Slide references retained by presentation metadata.
    SlideReferences,
    /// Custom presentation shows.
    CustomShows,
    /// Presentation sections.
    Sections,
    /// Document change lists.
    ChangeLists,
    /// Document changes.
    Changes,
    /// Presentation event records.
    Events,
    /// Animation records.
    Animations,
    /// Animation build records.
    AnimationBuilds,
    /// Animation paragraph templates.
    ParagraphTemplates,
    /// Animation time-filter points.
    TimeFilterPoints,
    /// Shapes or shape records.
    Shapes,
    /// Placeholder shape records in a newly authored slide.
    Placeholders,
    /// Shape-property replacement selections.
    ShapeProperties,
    /// Office Math equations discovered in a slide.
    Equations,
    /// Text-box shape records in a streaming slide.
    TextBoxes,
    /// Text replacement selections staged in one edit.
    TextReplacements,
    /// Ink annotation records.
    Annotations,
    /// Ink trace records.
    Traces,
    /// Ink trace groups.
    TraceGroups,
    /// Ink trace points.
    TracePoints,
    /// Ink action anchors.
    InkAnchors,
    /// Ink action picture children.
    PictureChildren,
    /// Picture records carrying media resources.
    MediaPictures,
    /// Media bookmark records.
    Bookmarks,
    /// Embedded chart series.
    ChartSeries,
    /// Embedded chart points.
    ChartPoints,
    /// Embedded controls.
    Controls,
    /// Embedded OLE objects.
    OleObjects,
    /// Content-part anchors.
    ContentParts,
    /// Table rows.
    TableRows,
    /// Table cells.
    TableCells,
    /// Table style definitions.
    TableStyles,
    /// Modern comment records and replies.
    CommentEntries,
    /// Modern comment replies.
    CommentReplies,
    /// Modern comment author records.
    Authors,
    /// Action setting records.
    ActionSettings,
    /// Shape designer tag records.
    DesignerTags,
    /// Text objects in a slide.
    TextObjects,
    /// Text runs in a slide.
    TextRuns,
    /// Zoom entries.
    ZoomEntries,
    /// Caption tracks.
    CaptionTracks,
    /// Caption text lines.
    CaptionLines,
    /// Caption text blocks.
    CaptionBlocks,
    /// Caption cue settings.
    CaptionSettings,
    /// Presentation tags.
    Tags,
    /// References or identities.
    References,
    /// Payload bytes.
    PayloadBytes,
    /// Serialized archive bytes.
    ArchiveBytes,
    /// Durable patch bytes.
    PatchBytes,
    /// Source bytes.
    SourceBytes,
    /// Retained bytes.
    RetainedBytes,
    /// A bounded index or position.
    Index,
    /// A bounded namespace byte sequence.
    NamespaceBytes,
    /// Namespace bindings or declarations.
    NamespaceBindings,
    /// Namespace prefixes.
    NamespacePrefixes,
    /// Extension count or extension list.
    Extensions,
    /// Extension payload bytes.
    ExtensionBytes,
    /// Aggregate Designer tag name and value bytes.
    DesignerTagBytes,
    /// Password/verifier bytes.
    PasswordBytes,
    /// Shape identifier values.
    ShapeIds,
    /// Slide identifier values.
    SlideIds,
    /// Snapshot entries or their bounded aggregate.
    Snapshots,
    /// Proof bytes used by a source-backed plan.
    ProofBytes,
    /// Bounded transient memory admission bytes.
    MemoryBytes,
    /// Embedded font count.
    Fonts,
    /// Presentation view guide entries.
    Guides,
    /// Presentation property recent-color entries.
    Colors,
    /// Model3D child entries.
    ModelEntries,
    /// One embedded font payload.
    FontBytes,
    /// Aggregate embedded-font payload bytes.
    TotalFontBytes,
}

impl CodecLimitKind {
    /// A content-free diagnostic label.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::XmlBytes => "XML bytes",
            Self::XmlDepth => "XML depth",
            Self::XmlNodes => "XML nodes",
            Self::XmlEvents => "XML events",
            Self::XmlAttributes => "XML attributes",
            Self::XmlAttributeBytes => "XML attribute bytes",
            Self::StringBytes => "string bytes",
            Self::TextBytes => "text bytes",
            Self::PartBytes => "part bytes",
            Self::Parts => "parts",
            Self::Relationships => "relationships",
            Self::RelationshipBytes => "relationship bytes",
            Self::ReferenceBytes => "reference bytes",
            Self::Names => "names",
            Self::NameBytes => "name bytes",
            Self::Slides => "slides",
            Self::SlideReferences => "slide references",
            Self::CustomShows => "custom shows",
            Self::Sections => "sections",
            Self::ChangeLists => "change lists",
            Self::Changes => "changes",
            Self::Events => "events",
            Self::Animations => "animations",
            Self::AnimationBuilds => "animation builds",
            Self::ParagraphTemplates => "paragraph templates",
            Self::TimeFilterPoints => "time-filter points",
            Self::Shapes => "shapes",
            Self::Placeholders => "placeholders",
            Self::ShapeProperties => "shape properties",
            Self::Equations => "equations",
            Self::TextBoxes => "text boxes",
            Self::TextReplacements => "text replacements",
            Self::Annotations => "annotations",
            Self::Traces => "traces",
            Self::TraceGroups => "trace groups",
            Self::TracePoints => "trace points",
            Self::InkAnchors => "ink anchors",
            Self::PictureChildren => "picture children",
            Self::MediaPictures => "media pictures",
            Self::Bookmarks => "bookmarks",
            Self::ChartSeries => "chart series",
            Self::ChartPoints => "chart points",
            Self::Controls => "controls",
            Self::OleObjects => "OLE objects",
            Self::ContentParts => "content parts",
            Self::TableRows => "table rows",
            Self::TableCells => "table cells",
            Self::TableStyles => "table styles",
            Self::CommentEntries => "comment entries",
            Self::CommentReplies => "comment replies",
            Self::Authors => "authors",
            Self::ActionSettings => "action settings",
            Self::DesignerTags => "designer tags",
            Self::TextObjects => "text objects",
            Self::TextRuns => "text runs",
            Self::ZoomEntries => "zoom entries",
            Self::CaptionTracks => "caption tracks",
            Self::CaptionLines => "caption lines",
            Self::CaptionBlocks => "caption blocks",
            Self::CaptionSettings => "caption settings",
            Self::Tags => "tags",
            Self::References => "references",
            Self::PayloadBytes => "payload bytes",
            Self::ArchiveBytes => "archive bytes",
            Self::PatchBytes => "patch bytes",
            Self::SourceBytes => "source bytes",
            Self::RetainedBytes => "retained bytes",
            Self::Index => "index",
            Self::NamespaceBytes => "namespace bytes",
            Self::NamespaceBindings => "namespace bindings",
            Self::NamespacePrefixes => "namespace prefixes",
            Self::Extensions => "extensions",
            Self::ExtensionBytes => "extension bytes",
            Self::DesignerTagBytes => "Designer tag bytes",
            Self::PasswordBytes => "password bytes",
            Self::ShapeIds => "shape identifiers",
            Self::SlideIds => "slide identifiers",
            Self::Snapshots => "snapshots",
            Self::ProofBytes => "proof bytes",
            Self::MemoryBytes => "memory bytes",
            Self::Fonts => "fonts",
            Self::Guides => "guides",
            Self::Colors => "colors",
            Self::ModelEntries => "model entries",
            Self::FontBytes => "font bytes",
            Self::TotalFontBytes => "total font bytes",
        }
    }
}

impl std::fmt::Display for CodecLimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Parts => formatter.write_str("changed parts"),
            Self::PatchBytes => formatter.write_str("durable patch bytes"),
            Self::TextBytes => formatter.write_str("replacement text bytes"),
            Self::HistoryEntries => formatter.write_str("history entries"),
            Self::HistoryBytes => formatter.write_str("history bytes"),
            Self::RetainedCandidateBytes => formatter.write_str("retained candidate bytes"),
            Self::RetainedMceBytes => formatter.write_str("retained markup-compatibility bytes"),
            Self::Codec(kind) => kind.fmt(formatter),
            Self::Common(kind) => kind.fmt(formatter),
        }
    }
}

/// Stable classification for a shape-transfer refusal that would otherwise
/// publish a lossy or dangling shape graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ShapeTransferRefusal {
    /// The selector names a grouped child instead of its transferable top-level owner.
    NestedShape,
    /// A `p:contentPart` has no common non-visual identity to return or remap.
    ContentPart,
    /// The selected closure contains an extension shape unknown to this release.
    UnknownExtensionShape,
    /// A generic `graphicFrame` payload has no typed dependency vocabulary.
    UnclassifiedGraphicFrame,
    /// A known common shape lacks its required non-visual identity.
    MissingIdentity,
    /// A connector endpoint does not resolve to a transferable shape on the source slide.
    UnresolvedConnectorEndpoint,
}

/// Stable classification for a whole-slide copy plan that cannot prove a
/// complete, independent dependency closure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SlideCopyRefusal {
    /// The source package contains OPC digital-signature infrastructure.
    SignedPackage,
    /// The presentation carries a modify-password verifier.
    ProtectedPresentation,
    /// Markup-compatibility input would require choosing or rewriting a branch.
    MarkupCompatibility,
    /// The slide or an owned XML dependency has an unmodeled semantic surface.
    UnknownSemanticSurface,
    /// A presentation-global or cross-slide owner is reachable from the slide.
    SharedOwner,
    /// A relationship family is outside the explicitly supported closure.
    UnsupportedRelationship,
    /// Owned dependency edges contain a cycle whose rewrite policy is not modeled.
    DependencyCycle,
    /// The package or selected slide has ambiguous physical ownership.
    AmbiguousTopology,
    /// A DrawingML table depends on the presentation-global table-style catalog.
    GlobalTableStyle,
    /// The physical package carries an unknown non-Part ZIP member that this
    /// topology-changing operation cannot preserve or authorize.
    UnknownPhysicalMember,
    /// The destination holds a caller-defined `Part` implementation, and the
    /// copy frames source-compressed media, so it could only publish a reopen
    /// of the destination's bytes, in which that part would become a built-in
    /// one. Planning the copy against this destination records the
    /// recompressing route, which keeps the part.
    CallerDefinedPart,
}

/// Durable `PresentationML` patch families, each naming the revision proof its
/// serialized header carries.
///
/// A durable patch embeds complete-package revisions computed by one exact
/// algorithm. Change 0655 redefined that algorithm (`litchi-pptx-opened-v1` →
/// `litchi-pptx-opened-v2`), so the two durable families carry a new magic and
/// a patch serialized under the superseded magic is refused by name rather
/// than compared across algebras. Change 0742 made a cross-presentation copy's
/// serialized-archive (physical) revisions depend on how its candidate
/// encodes copied image members, which the header now records; `LPCP0003`
/// carries no such field, so it is superseded by `LPCP0004` and refused by
/// name rather than read implicitly as either encoding. See
/// [`Error::PatchConflict`].
///
/// This enum never carries attacker-supplied bytes: an unrecognized magic is
/// not classified here at all and keeps the existing
/// [`Error::Invalid`] refusal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DurablePatchFormat {
    /// `LPRM0001`: slide removal, superseded `litchi-pptx-opened-v1` proof.
    SlideRemovalV1,
    /// `LPRM0002`: slide removal, current `litchi-pptx-opened-v2` proof.
    SlideRemovalV2,
    /// `LPCP0002`: cross-presentation slide copy, superseded semantic proof.
    CrossSlideCopyV2,
    /// `LPCP0003`: cross-presentation slide copy without the recorded
    /// copied-media encoding its physical revisions now depend on
    /// (superseded by change 0742).
    CrossSlideCopyV3,
    /// `LPCP0004`: cross-presentation slide copy, current semantic and
    /// physical proofs, with the candidate's copied-media encoding recorded.
    CrossSlideCopyV4,
}

impl DurablePatchFormat {
    /// Eight-byte durable header magic that identifies this format.
    #[must_use]
    pub const fn magic(self) -> &'static [u8; 8] {
        match self {
            Self::SlideRemovalV1 => b"LPRM0001",
            Self::SlideRemovalV2 => b"LPRM0002",
            Self::CrossSlideCopyV2 => b"LPCP0002",
            Self::CrossSlideCopyV3 => b"LPCP0003",
            Self::CrossSlideCopyV4 => b"LPCP0004",
        }
    }
}

impl std::fmt::Display for DurablePatchFormat {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::SlideRemovalV1 => "LPRM0001",
            Self::SlideRemovalV2 => "LPRM0002",
            Self::CrossSlideCopyV2 => "LPCP0002",
            Self::CrossSlideCopyV3 => "LPCP0003",
            Self::CrossSlideCopyV4 => "LPCP0004",
        })
    }
}

/// Stable classification for a whole-slide removal plan that cannot prove a
/// complete, independent deletion boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SlideRemovalRefusal {
    /// The source package contains OPC digital-signature infrastructure.
    SignedPackage,
    /// The package is macro-enabled or contains VBA infrastructure.
    MacroEnabledPackage,
    /// The presentation carries a modify-password verifier.
    ProtectedPresentation,
    /// Markup-compatibility input would require choosing or rewriting a branch.
    MarkupCompatibility,
    /// The selected slide or presentation owner has an unmodeled semantic surface.
    UnknownSemanticSurface,
    /// Another package resource or presentation feature references the selected slide.
    SharedOwner,
    /// The selected slide has a relationship outside the dependency-free boundary.
    UnsupportedRelationship,
    /// The package or selected slide has ambiguous physical ownership.
    AmbiguousTopology,
    /// The physical package carries an unknown non-Part ZIP member that this
    /// topology-changing operation cannot preserve.
    UnknownPhysicalMember,
    /// A presentation must retain at least one slide.
    FinalSlide,
}

/// Stable classification for a source-backed slide-order edit that cannot
/// prove every positional owner remains coherent after the move.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SlideOrderRefusal {
    /// Markup-compatibility input would require choosing or rewriting a branch.
    MarkupCompatibility,
    /// PowerPoint section metadata would retain the old slide positions.
    Sections,
    /// The presentation view-properties part owns positional state.
    ViewProperties,
    /// The presentation-properties part owns positional state.
    PresentationProperties,
    /// Another known positional semantic owner is present.
    PositionalOwner,
    /// The raw slide list or its `(id, r:id)` bindings are ambiguous.
    AmbiguousBindings,
    /// A direct presentation-root modify verifier is not safe to preserve through reordering.
    ModifyVerifier,
}

impl SlideCopyRefusal {
    /// The refusal kind litchi-ops classifies this reason by (record 0351).
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SignedPackage | Self::ProtectedPresentation => RefusalKind::Protected,
            Self::CallerDefinedPart => RefusalKind::UnsupportedOperation,
            Self::MarkupCompatibility
            | Self::UnknownSemanticSurface
            | Self::SharedOwner
            | Self::UnsupportedRelationship
            | Self::DependencyCycle
            | Self::AmbiguousTopology
            | Self::GlobalTableStyle
            | Self::UnknownPhysicalMember => RefusalKind::UnsupportedContent,
        }
    }
}

impl std::fmt::Display for SlideCopyRefusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::SignedPackage => "signed package",
            Self::ProtectedPresentation => "protected presentation",
            Self::MarkupCompatibility => "markup-compatibility surface",
            Self::UnknownSemanticSurface => "unknown semantic surface",
            Self::SharedOwner => "shared or cross-slide owner",
            Self::UnsupportedRelationship => "unsupported relationship family",
            Self::DependencyCycle => "dependency cycle",
            Self::AmbiguousTopology => "ambiguous package topology",
            Self::GlobalTableStyle => "presentation-global table style",
            Self::UnknownPhysicalMember => "unknown physical package member",
            Self::CallerDefinedPart => "caller-defined destination part",
        })
    }
}

impl SlideRemovalRefusal {
    /// The refusal kind litchi-ops classifies this reason by (record 0351).
    ///
    /// [`Self::FinalSlide`] is an invalid value: the selection would leave
    /// the presentation without slides. litchi-ops reports it as
    /// `invalid_value`, by the shared refusal-kind contract.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SignedPackage | Self::ProtectedPresentation => RefusalKind::Protected,
            Self::FinalSlide => RefusalKind::InvalidValue,
            Self::MacroEnabledPackage
            | Self::MarkupCompatibility
            | Self::UnknownSemanticSurface
            | Self::SharedOwner
            | Self::UnsupportedRelationship
            | Self::AmbiguousTopology
            | Self::UnknownPhysicalMember => RefusalKind::UnsupportedContent,
        }
    }
}

impl std::fmt::Display for SlideRemovalRefusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::SignedPackage => "signed package",
            Self::MacroEnabledPackage => "macro-enabled package",
            Self::ProtectedPresentation => "protected presentation",
            Self::MarkupCompatibility => "markup-compatibility surface",
            Self::UnknownSemanticSurface => "unknown semantic surface",
            Self::SharedOwner => "shared or cross-slide owner",
            Self::UnsupportedRelationship => "unsupported relationship family",
            Self::AmbiguousTopology => "ambiguous package topology",
            Self::UnknownPhysicalMember => "unknown physical package member",
            Self::FinalSlide => "final remaining slide",
        })
    }
}

impl SlideOrderRefusal {
    /// The refusal kind litchi-ops classifies this reason by (record 0351).
    ///
    /// [`Self::AmbiguousBindings`] is unsupported content; litchi-ops
    /// reports it as `corrupt_file`, because ambiguous slide bindings are a
    /// malformed source rather than a route limitation.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::ModifyVerifier => RefusalKind::Protected,
            Self::MarkupCompatibility
            | Self::Sections
            | Self::ViewProperties
            | Self::PresentationProperties
            | Self::PositionalOwner
            | Self::AmbiguousBindings => RefusalKind::UnsupportedContent,
        }
    }
}

impl std::fmt::Display for SlideOrderRefusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::MarkupCompatibility => "markup-compatibility surface",
            Self::Sections => "presentation sections",
            Self::ViewProperties => "presentation view properties",
            Self::PresentationProperties => "presentation properties",
            Self::PositionalOwner => "positional semantic owner",
            Self::AmbiguousBindings => "ambiguous slide-order bindings",
            Self::ModifyVerifier => "presentation modify verifier",
        })
    }
}

impl ShapeTransferRefusal {
    /// The refusal kind litchi-ops classifies this reason by (record 0351).
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NestedShape | Self::MissingIdentity | Self::UnresolvedConnectorEndpoint => {
                RefusalKind::UnsupportedTarget
            },
            Self::ContentPart | Self::UnknownExtensionShape | Self::UnclassifiedGraphicFrame => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl std::fmt::Display for ShapeTransferRefusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::NestedShape => "nested shape selection",
            Self::ContentPart => "content-part shape",
            Self::UnknownExtensionShape => "unknown extension shape",
            Self::UnclassifiedGraphicFrame => "unclassified graphic frame",
            Self::MissingIdentity => "missing non-visual identity",
            Self::UnresolvedConnectorEndpoint => "unresolved connector endpoint",
        })
    }
}

/// Failure to decode or encode a `PresentationML` capability.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// Filesystem access or atomic output replacement failed.
    #[error("PPTX I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// The underlying OPC graph is malformed or could not be read safely.
    #[error("PPTX OPC error: {0}")]
    Opc(#[from] litchi_opc::OpcError),

    /// The XML stream is not well formed or cannot be decoded safely.
    #[error("invalid PresentationML XML: {0}")]
    Xml(String),

    /// The document violates a `PresentationML` structural or value invariant.
    #[error("invalid PresentationML: {0}")]
    Invalid(String),

    /// A bounded PresentationML resource was exhausted.
    ///
    /// `observed` is the measured charge.  Arithmetic overflow is represented
    /// by `usize::MAX`, never by a fabricated `maximum + 1` value.
    #[error("PresentationML {kind} exceeds the limit of {maximum} (observed {observed})")]
    Limit {
        /// Dimension that exceeded its configured limit.
        kind: LimitKind,
        /// Measured charge, or `usize::MAX` when the charge overflowed.
        observed: usize,
        /// Active upper bound.
        maximum: usize,
    },

    /// A caller-selected PresentationML limit was invalid.
    #[error("invalid PresentationML {kind} limit: {value} (maximum {maximum})")]
    InvalidLimit {
        /// Dimension whose value was invalid.
        kind: LimitKind,
        /// Supplied value.
        value: usize,
        /// Hard maximum permitted for this dimension.
        maximum: usize,
    },

    /// A bounded `PresentationML` operation could not reserve required memory.
    #[error("could not reserve memory for PresentationML {resource}: {source}")]
    Allocation {
        /// Resource whose bounded plan could not be reserved.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
    },

    /// A bounded structural byte count overflowed before publication.
    ///
    /// This carries no fabricated observed or configured value: the
    /// operation could not represent the count in its structural type.
    #[error("PPTX {resource} capacity overflow")]
    CapacityOverflow {
        /// Structural byte count that overflowed.
        resource: &'static str,
    },

    /// A related part has a content type forbidden by `PresentationML`.
    #[error("invalid PPTX content type: expected {expected}, got {actual}")]
    ContentType {
        /// Required content type.
        expected: String,
        /// Content type found in the package.
        actual: String,
    },

    /// A `PresentationML` relationship is missing or malformed.
    #[error("invalid PresentationML relationship: {0}")]
    Relationship(String),

    /// A relationship target part is absent from the OPC package.
    #[error("PresentationML part not found: {0}")]
    PartNotFound(String),

    /// No tag has the requested semantic name.
    #[error("tag name '{0}' was not found")]
    NameNotFound(String),

    /// A numeric tag selector is outside the checked list bounds.
    #[error("tag index {index} is outside a list of length {len}")]
    IndexOutOfBounds {
        /// Requested zero-based index.
        index: usize,
        /// Current list length.
        len: usize,
    },

    /// Malformed producer input contains multiple caseless-equivalent names.
    #[error("tag name '{name}' is ambiguous ({matches} matches)")]
    AmbiguousName {
        /// Selector spelling supplied by the caller.
        name: String,
        /// Number of matching tags.
        matches: usize,
    },

    /// No slide has the requested semantic name.
    #[error("slide name '{0}' was not found")]
    SlideNameNotFound(String),

    /// More than one slide has the requested semantic name.
    #[error("slide name '{name}' is ambiguous ({matches} matches)")]
    AmbiguousSlideName {
        /// Selector spelling supplied by the caller.
        name: String,
        /// Number of matching slides.
        matches: usize,
    },

    /// No registered slide layout has the requested exact name.
    #[error("slide layout name '{0}' was not found")]
    SlideLayoutNameNotFound(String),

    /// Multiple registered slide layouts have the requested exact name.
    #[error("slide layout name '{name}' is ambiguous ({matches} matches)")]
    AmbiguousSlideLayoutName {
        /// Selector spelling supplied by the caller.
        name: String,
        /// Number of matching layouts.
        matches: usize,
    },

    /// A numeric slide selector is outside the checked presentation bounds.
    #[error("slide index {index} is outside a presentation of length {len}")]
    SlideIndexOutOfBounds {
        /// Requested zero-based index.
        index: usize,
        /// Current slide count.
        len: usize,
    },

    /// A mutation would create a caseless-equivalent duplicate name.
    #[error("tag name '{name}' conflicts with {matches} existing tag(s)")]
    DuplicateName {
        /// Spelling supplied by the caller.
        name: String,
        /// Number of conflicting tags.
        matches: usize,
    },

    /// A requested reorder is not a complete list permutation.
    #[error("tag reorder has {actual} selectors; expected {expected}")]
    OrderLength {
        /// Required selector count.
        expected: usize,
        /// Supplied selector count.
        actual: usize,
    },

    /// A requested reorder selects the same physical tag more than once.
    #[error("tag reorder selects index {index} more than once")]
    DuplicateSelection {
        /// Repeated physical index.
        index: usize,
    },

    /// No embedded font has the requested semantic typeface.
    #[error("embedded font '{0}' was not found")]
    FontNotFound(String),

    /// Malformed producer input contains multiple caseless-equivalent typefaces.
    #[error("embedded font '{name}' is ambiguous ({matches} matches)")]
    AmbiguousFontName {
        /// Selector spelling supplied by the caller.
        name: String,
        /// Number of matching physical entries.
        matches: usize,
    },

    /// A numeric embedded-font selector is outside the checked list bounds.
    #[error("embedded-font index {index} is outside a list of length {len}")]
    FontIndexOutOfBounds {
        /// Requested zero-based index.
        index: usize,
        /// Current embedded-font count.
        len: usize,
    },

    /// A mutation would create a Unicode-caseless duplicate typeface.
    #[error("embedded font '{name}' conflicts with {matches} existing font(s)")]
    DuplicateFontName {
        /// Typeface spelling supplied by the caller.
        name: String,
        /// Number of conflicting fonts.
        matches: usize,
    },

    /// A complete embedded-font reorder has the wrong number of selectors.
    #[error("embedded-font reorder has {actual} selectors; expected {expected}")]
    FontOrderLength {
        /// Required selector count.
        expected: usize,
        /// Supplied selector count.
        actual: usize,
    },

    /// A reorder selects the same physical embedded font more than once.
    #[error("embedded-font reorder selects index {index} more than once")]
    DuplicateFontSelection {
        /// Repeated physical index.
        index: usize,
    },

    /// A package URI supplied to the `PresentationML` graph is invalid.
    #[error("invalid PresentationML URI: {0}")]
    Uri(String),

    /// A well-formed request this route refuses (record 0351): the kind
    /// classifies it, the operation and reason are static text.
    #[error("PPTX {0}")]
    Refused(litchi_core::edit::Refusal),

    /// A patch or staged plan does not match the package it is applied to:
    /// it was made against different exact source bytes or another revision
    /// (record 0351).
    #[error("PPTX patch base does not match the target package")]
    PatchConflict,

    /// A changed package still carries signature infrastructure.
    #[error("PPTX signed package requires an explicit signature disposition")]
    Signed,

    /// A source-backed package changed while the operation was in progress.
    #[error("PPTX source changed while the operation was in progress")]
    SourceChanged {
        /// Version captured before the operation.
        expected: litchi_core::SourceVersion,
        /// Version observed when the operation completed.
        observed: litchi_core::SourceVersion,
    },

    /// A broken internal invariant: state the crate itself should have
    /// provided is missing or inconsistent.
    #[error("PPTX {operation} internal invariant failed: {reason}")]
    Internal {
        /// Operation that detected the broken invariant.
        operation: &'static str,
        /// The invariant that failed.
        reason: &'static str,
    },

    /// An owning ingress route was given a policy only a positional-source
    /// owner can honor.
    #[error("owning PPTX ingress does not support the {policy} open policy")]
    UnsupportedOpenPolicy {
        /// The refused policy.
        policy: crate::OpenPolicy,
    },

    /// Managed package encryption or decryption failed.
    #[cfg(feature = "encryption")]
    #[error("PPTX package encryption error: {0}")]
    Encryption(#[from] litchi_crypto::ooxml::Error),

    /// An operation would violate retained package-encryption policy.
    #[cfg(feature = "encryption")]
    #[error("PPTX encryption policy rejected {operation}: {source}")]
    EncryptionPolicy {
        /// Operation rejected by the policy.
        operation: &'static str,
        /// Host-independent retained-provenance failure.
        #[source]
        source: litchi_ooxml_common::package_encryption::PolicyError,
    },

    /// A semantic shape selector is missing, ambiguous, or outside checked bounds.
    #[error("PresentationML shape lookup error: {0}")]
    ShapeLookup(#[from] crate::shape::LookupError),

    /// Markup-compatibility processing failed.
    #[error("PresentationML markup compatibility error: {0}")]
    MarkupCompatibility(#[from] litchi_ooxml_common::mce::Error),

    /// Shared OOXML attribute decoding failed.
    #[error("PresentationML attribute decoding error: {0}")]
    Decode(#[from] litchi_ooxml_common::XmlError),

    /// A shared `DrawingML` chart or primitive codec failed.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::Error),

    /// Shared custom document-property graph or value validation failed.
    #[error("PPTX custom document-property error: {0}")]
    CustomProperties(#[source] litchi_ooxml_common::Error),

    /// Optional system-font discovery, subsetting, or EOT preparation failed.
    #[cfg(feature = "automatic-fonts")]
    #[error("PPTX font embedding error: {0}")]
    Font(#[from] litchi_fonts::FontError),

    /// Writing into the requested text sink failed.
    #[error("could not encode PresentationML text")]
    Write,
}

impl Error {
    /// A refusal of `operation` for `reason`, classified by `kind`.
    pub(crate) const fn refused(
        kind: litchi_core::edit::RefusalKind,
        operation: &'static str,
        reason: &'static str,
    ) -> Self {
        Self::Refused(litchi_core::edit::Refusal::new(kind, operation, reason))
    }

    /// A broken internal invariant.
    pub(crate) const fn internal(operation: &'static str, reason: &'static str) -> Self {
        Self::Internal { operation, reason }
    }

    /// A whole-slide removal refusal with a stable operation and reason.
    pub(crate) const fn slide_removal_refused(kind: SlideRemovalRefusal) -> Self {
        let reason = match kind {
            SlideRemovalRefusal::SignedPackage => "the package is signed",
            SlideRemovalRefusal::MacroEnabledPackage => {
                "macro infrastructure is outside dependency-free slide removal"
            },
            SlideRemovalRefusal::ProtectedPresentation => {
                "the presentation has an active modify-password verifier"
            },
            SlideRemovalRefusal::MarkupCompatibility => {
                "slide removal cannot preserve markup-compatibility content"
            },
            SlideRemovalRefusal::UnknownSemanticSurface => {
                "slide removal cannot preserve an unsupported semantic surface"
            },
            SlideRemovalRefusal::SharedOwner => "the selected slide has a shared owner",
            SlideRemovalRefusal::UnsupportedRelationship => {
                "the selected slide has an unsupported relationship"
            },
            SlideRemovalRefusal::AmbiguousTopology => "the selected slide has ambiguous topology",
            SlideRemovalRefusal::UnknownPhysicalMember => {
                "the package contains an unknown physical member"
            },
            SlideRemovalRefusal::FinalSlide => "a presentation must retain at least one slide",
        };
        Self::refused(kind.kind(), "slide_removal", reason)
    }

    /// A source-backed slide-order refusal with a stable operation and reason.
    /// Ambiguous bindings are malformed source data, so they retain the
    /// corrupt-file route rather than becoming a well-formed refusal.
    pub(crate) fn slide_order_refused(kind: SlideOrderRefusal) -> Self {
        match kind {
            SlideOrderRefusal::AmbiguousBindings => {
                Self::Invalid("slide-order bindings are ambiguous".to_owned())
            },
            _ => Self::refused(
                kind.kind(),
                "slide_order",
                match kind {
                    SlideOrderRefusal::ModifyVerifier => {
                        "the presentation has an active modify-password verifier"
                    },
                    SlideOrderRefusal::MarkupCompatibility => {
                        "slide order cannot preserve markup-compatibility content"
                    },
                    SlideOrderRefusal::Sections => {
                        "slide order cannot preserve presentation sections"
                    },
                    SlideOrderRefusal::ViewProperties => {
                        "slide order cannot preserve presentation view properties"
                    },
                    SlideOrderRefusal::PresentationProperties => {
                        "slide order cannot preserve presentation properties"
                    },
                    SlideOrderRefusal::PositionalOwner => {
                        "slide order cannot preserve positional metadata"
                    },
                    SlideOrderRefusal::AmbiguousBindings => "slide-order bindings are ambiguous",
                },
            ),
        }
    }

    /// A whole-slide copy refusal with a stable operation and reason.
    pub(crate) const fn slide_copy_refused(kind: SlideCopyRefusal) -> Self {
        let reason = match kind {
            SlideCopyRefusal::SignedPackage => {
                "the selected slide cannot be copied losslessly because the package is signed"
            },
            SlideCopyRefusal::ProtectedPresentation => {
                "the selected slide cannot be copied losslessly because the presentation is protected"
            },
            SlideCopyRefusal::MarkupCompatibility => {
                "the selected slide cannot be copied losslessly because it contains markup-compatibility content"
            },
            SlideCopyRefusal::UnknownSemanticSurface => {
                "the selected slide cannot be copied losslessly because its semantic surface is unsupported"
            },
            SlideCopyRefusal::SharedOwner => {
                "the selected slide cannot be copied losslessly because it has a shared owner"
            },
            SlideCopyRefusal::UnsupportedRelationship => {
                "the selected slide cannot be copied losslessly because it has an unsupported relationship"
            },
            SlideCopyRefusal::DependencyCycle => {
                "the selected slide cannot be copied losslessly because its dependency graph has a cycle"
            },
            SlideCopyRefusal::AmbiguousTopology => {
                "the selected slide cannot be copied losslessly because its topology is ambiguous"
            },
            SlideCopyRefusal::GlobalTableStyle => {
                "the selected slide cannot be copied losslessly because it uses a presentation-global table style"
            },
            SlideCopyRefusal::UnknownPhysicalMember => {
                "the selected slide cannot be copied losslessly because the package has an unknown physical member"
            },
            SlideCopyRefusal::CallerDefinedPart => {
                "the selected slide cannot be copied losslessly because the destination has caller-defined parts"
            },
        };
        Self::refused(kind.kind(), "slide_copy", reason)
    }

    /// A shape-transfer refusal with a stable operation and reason.
    pub(crate) const fn shape_transfer_refused(kind: ShapeTransferRefusal) -> Self {
        let reason = match kind {
            ShapeTransferRefusal::NestedShape => {
                "the selected shape is nested inside another shape"
            },
            ShapeTransferRefusal::ContentPart => "the selected shape is a content part",
            ShapeTransferRefusal::UnknownExtensionShape => {
                "the selected shape has an unknown extension"
            },
            ShapeTransferRefusal::UnclassifiedGraphicFrame => {
                "the selected graphic frame is not classified"
            },
            ShapeTransferRefusal::MissingIdentity => {
                "the selected shape has no non-visual identity"
            },
            ShapeTransferRefusal::UnresolvedConnectorEndpoint => {
                "the selected connector endpoint is unresolved"
            },
        };
        Self::refused(kind.kind(), "shape_transfer", reason)
    }

    /// The refusal this error carries, when it is one.
    #[must_use]
    pub const fn refusal(&self) -> Option<&litchi_core::edit::Refusal> {
        match self {
            Self::Refused(refusal) => Some(refusal),
            _ => None,
        }
    }
}

impl From<litchi_ooxml_common::Error> for Error {
    fn from(error: litchi_ooxml_common::Error) -> Self {
        match error {
            litchi_ooxml_common::Error::Refused(refusal) => Self::Refused(refusal),
            litchi_ooxml_common::Error::PatchConflict => Self::PatchConflict,
            litchi_ooxml_common::Error::Signed => Self::Signed,
            litchi_ooxml_common::Error::SourceChanged { expected, observed } => {
                Self::SourceChanged { expected, observed }
            },
            litchi_ooxml_common::Error::Internal(reason) => Self::internal("shared OOXML", reason),
            other => Self::CustomProperties(other),
        }
    }
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::{CodecLimitKind, Error, LimitKind};

    #[test]
    fn typed_limit_keeps_dimension_and_exact_observation() {
        let error = Error::Limit {
            kind: LimitKind::Codec(CodecLimitKind::TracePoints),
            observed: 17,
            maximum: 16,
        };
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Codec(CodecLimitKind::TracePoints),
                observed: 17,
                maximum: 16,
            }
        ));
        assert_eq!(
            error.to_string(),
            "PresentationML trace points exceeds the limit of 16 (observed 17)"
        );
    }

    #[test]
    fn invalid_limit_reports_the_supplied_value_without_resource_strings() {
        let error = Error::InvalidLimit {
            kind: LimitKind::Codec(CodecLimitKind::Annotations),
            value: 0,
            maximum: 4096,
        };
        assert!(matches!(
            error,
            Error::InvalidLimit {
                kind: LimitKind::Codec(CodecLimitKind::Annotations),
                value: 0,
                maximum: 4096,
            }
        ));
        assert_eq!(
            error.to_string(),
            "invalid PresentationML annotations limit: 0 (maximum 4096)"
        );
    }
}
