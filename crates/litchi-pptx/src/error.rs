//! Typed PPTX failures.

use std::collections::TryReserveError;

use thiserror::Error;

/// Result of a PPTX operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Failure to decode or encode a PresentationML capability.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The underlying OPC graph is malformed or could not be read safely.
    #[error("PPTX OPC error: {0}")]
    Opc(#[from] litchi_opc::OpcError),

    /// The XML stream is not well formed or cannot be decoded safely.
    #[error("invalid PresentationML XML: {0}")]
    Xml(String),

    /// The document violates a PresentationML structural or value invariant.
    #[error("invalid PresentationML: {0}")]
    Invalid(String),

    /// A bounded decoder resource was exhausted.
    #[error("PresentationML {resource} exceeds the limit of {limit}")]
    Limit {
        /// Resource that exceeded its configured limit.
        resource: &'static str,
        /// Active upper bound.
        limit: usize,
    },

    /// A bounded PresentationML operation could not reserve required memory.
    #[error("could not reserve memory for PresentationML {resource}: {source}")]
    Allocation {
        /// Resource whose bounded plan could not be reserved.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
    },

    /// A related part has a content type forbidden by PresentationML.
    #[error("invalid PPTX content type: expected {expected}, got {actual}")]
    ContentType {
        /// Required content type.
        expected: String,
        /// Content type found in the package.
        actual: String,
    },

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

    /// A semantic shape selector is missing, ambiguous, or outside checked bounds.
    #[error("PresentationML shape lookup error: {0}")]
    ShapeLookup(#[from] crate::shape::LookupError),

    /// Markup-compatibility processing failed.
    #[error("PresentationML markup compatibility error: {0}")]
    MarkupCompatibility(#[from] litchi_ooxml_common::MceError),

    /// Shared OOXML attribute decoding failed.
    #[error("PresentationML attribute decoding error: {0}")]
    Decode(#[from] litchi_ooxml_common::XmlError),

    /// Writing into the requested text sink failed.
    #[error("could not encode PresentationML text")]
    Write,
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}
