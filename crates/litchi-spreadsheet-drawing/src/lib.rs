//! Shared `SpreadsheetML` drawing primitives.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (record 0770, workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]

pub mod chart;
pub mod objects;
pub mod shape;
pub mod vml;

use thiserror::Error;

/// Result of a shared `SpreadsheetML` drawing operation.
pub type Result<T> = std::result::Result<T, Error>;

/// A bounded resource dimension used by the shared SpreadsheetDrawing
/// scanners and rewriters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Top-level objects of one drawing.
    Elements,
    /// Bytes in one drawing part or VML member.
    DrawingBytes,
    /// Bytes in one staged image.
    ImageBytes,
    /// Anchors retained by the shape inventory.
    Anchors,
    /// Objects retained by the shape inventory, including group children.
    Objects,
    /// Nesting depth of shape groups.
    GroupDepth,
    /// XML nesting depth.
    XmlDepth,
    /// Retained shape text bytes.
    TextBytes,
    /// Retained text XML elements.
    TextElements,
    /// VML shapes retained by one scan.
    VmlShapes,
}

impl LimitKind {
    /// The hard safety ceiling for this dimension.
    pub const fn maximum(self) -> usize {
        match self {
            Self::Elements => objects::MAX_ELEMENTS,
            Self::DrawingBytes => objects::MAX_DRAWING_BYTES,
            Self::ImageBytes => objects::MAX_IMAGE_BYTES,
            Self::Anchors => shape::MAX_ANCHORS_PER_DRAWING,
            Self::Objects => shape::MAX_OBJECTS_PER_DRAWING,
            Self::GroupDepth => shape::MAX_GROUP_DEPTH,
            Self::XmlDepth => shape::MAX_XML_DEPTH,
            Self::TextBytes => shape::MAX_TEXT_BYTES,
            Self::TextElements => shape::MAX_TEXT_ELEMENTS,
            Self::VmlShapes => vml::MAX_SHAPES,
        }
    }

    /// A stable, content-free description for diagnostics.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Elements => "worksheet drawing objects",
            Self::DrawingBytes => "worksheet drawing bytes",
            Self::ImageBytes => "worksheet image bytes",
            Self::Anchors => "SpreadsheetDrawing anchors",
            Self::Objects => "SpreadsheetDrawing objects",
            Self::GroupDepth => "SpreadsheetDrawing group depth",
            Self::XmlDepth => "SpreadsheetDrawing XML depth",
            Self::TextBytes => "SpreadsheetDrawing text bytes",
            Self::TextElements => "SpreadsheetDrawing text elements",
            Self::VmlShapes => "VML shapes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Failure while reading, validating, or writing shared `SpreadsheetML` drawing data.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The supplied chart structure violates an OOXML invariant.
    #[error("invalid SpreadsheetML drawing structure: {0}")]
    Invalid(String),
    /// A caller-selected resource budget was exceeded.
    #[error(
        "SpreadsheetML drawing {kind:?} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource dimension that was exceeded.
        kind: LimitKind,
        /// First observed value beyond the selected bound.
        observed: usize,
        /// Active upper bound.
        maximum: usize,
    },
    /// A caller supplied a zero or physically unsafe resource budget.
    #[error("invalid SpreadsheetML drawing {kind:?} limit {value} (allowed 1..={maximum})")]
    InvalidLimit {
        /// Resource dimension being configured.
        kind: LimitKind,
        /// Rejected configured value.
        value: usize,
        /// Hard safety ceiling.
        maximum: usize,
    },
    /// XML or byte encoding could not be decoded or emitted.
    #[error("SpreadsheetML drawing encoding error: {0}")]
    Encoding(String),
    /// Markup-compatibility preprocessing failed while reading drawing XML.
    #[error("SpreadsheetML drawing markup compatibility error: {0}")]
    Mce(#[from] litchi_ooxml_common::mce::Error),
    /// Shared OOXML attribute or character-reference decoding failed.
    #[error("SpreadsheetML drawing XML error: {0}")]
    Xml(#[from] litchi_ooxml_common::XmlError),
    /// A package part or relationship could not be resolved.
    #[error("SpreadsheetML drawing package error: {0}")]
    Opc(#[from] litchi_opc::OpcError),
    /// `DrawingML` chart decoding or validation failed.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::Error),
}

impl Error {
    /// Construct a typed observed-limit error.
    pub(crate) const fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Self {
        Self::Limit {
            kind,
            observed,
            maximum,
        }
    }

    /// Construct a typed invalid-policy error.
    pub(crate) const fn invalid_limit(kind: LimitKind, value: usize, maximum: usize) -> Self {
        Self::InvalidLimit {
            kind,
            value,
            maximum,
        }
    }
}
