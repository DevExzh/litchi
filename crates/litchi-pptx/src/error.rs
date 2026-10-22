//! Typed PPTX failures.

use std::collections::TryReserveError;

use thiserror::Error;

/// Result of a PPTX operation.
pub type Result<T> = std::result::Result<T, Error>;

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
    /// `invalid_argument`, the code every format uses for deleting the final
    /// slide.
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

    /// A bounded codec resource was exhausted.
    #[error("PresentationML {resource} exceeds the limit of {limit}")]
    Limit {
        /// Resource that exceeded its configured limit.
        resource: &'static str,
        /// Active upper bound.
        limit: usize,
    },

    /// A primary presentation transaction exceeded one dimension of its
    /// [`crate::opened::Limits`].
    #[error("PresentationML edit exceeds its {kind} limit of {maximum}")]
    EditLimit {
        /// The refused dimension.
        kind: crate::opened::LimitKind,
        /// Active upper bound.
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
    CustomProperties(#[from] litchi_ooxml_common::Error),

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

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}
