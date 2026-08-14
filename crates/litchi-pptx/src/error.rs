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
        })
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
    /// A whole-slide removal prerequisite was classified and refused without mutation.
    #[error("unsupported PresentationML slide removal plan ({kind}): {detail}")]
    SlideRemovalPlan {
        /// Machine-readable refusal family.
        kind: SlideRemovalRefusal,
        /// Bounded human-readable source context.
        detail: String,
    },

    /// A whole-slide copy prerequisite was classified and refused without mutation.
    #[error("unsupported PresentationML slide copy plan ({kind}): {detail}")]
    SlideCopyPlan {
        /// Machine-readable refusal family.
        kind: SlideCopyRefusal,
        /// Bounded human-readable source context.
        detail: String,
    },

    /// A common-shape transfer was classified and refused before publication.
    #[error("unsupported PresentationML shape transfer ({kind}): {detail}")]
    ShapeTransfer {
        /// Machine-readable refusal family.
        kind: ShapeTransferRefusal,
        /// Bounded human-readable source context.
        detail: String,
    },

    /// One atomic shape-text batch selected the same semantic shape twice.
    #[error("shape-text batch selects shape index {index} more than once")]
    DuplicateShapeTextSelection {
        /// Repeated zero-based pre-order shape index.
        index: usize,
    },

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

    /// A bounded `PresentationML` operation could not reserve required memory.
    #[error("could not reserve memory for PresentationML {resource}: {source}")]
    Allocation {
        /// Resource whose bounded plan could not be reserved.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
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

    /// A lossless facade operation cannot safely expose or replace a graph.
    #[error("unsafe PPTX edit during {operation}: {reason}")]
    UnsafeEdit {
        /// Operation that requested the edit.
        operation: &'static str,
        /// Reason the bounded facade rejected it.
        reason: &'static str,
    },

    /// A source-backed slide patch was created against different exact slide bytes.
    #[error("source-backed PPTX slide patch source is stale")]
    StaleSource,

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

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}
