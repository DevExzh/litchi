//! Shared `DrawingML` grammar independent of DOCX, PPTX, XLSX, and XLSB.
//!
//! [`chart`] owns the chart model and XML codec, while [`diagram`] owns the
//! host-neutral `SmartArt` model and part grammar. Concrete formats retain only
//! their package relationships and anchoring semantics.

#![forbid(unsafe_code)]
// Public DrawingML types intentionally mirror ECMA-376 names, builder
// ownership, and boolean-rich schema records. Renaming or reshaping them here
// would break every host format without changing wire correctness.
#![allow(
    clippy::double_must_use,
    clippy::missing_errors_doc,
    clippy::missing_panics_doc,
    clippy::module_name_repetitions,
    clippy::needless_pass_by_value,
    clippy::ref_option,
    clippy::return_self_not_must_use,
    clippy::struct_excessive_bools,
    reason = "shared DrawingML APIs retain specification-shaped names, ownership, and centrally documented failure contracts"
)]
// Streaming chart, theme, geometry, and transform parsers reuse event-local
// names as raw XML becomes typed vocabulary.
#![allow(
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    reason = "short-lived parser bindings track successive XML events and semantic projections"
)]
// Fixed-width conversions and invariant accesses are bounded by the surrounding
// schema validators; unknown XML events remain inert for forward compatibility.
#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_sign_loss,
    clippy::expect_used,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::wildcard_enum_match_arm,
    reason = "validated DrawingML wire bounds and parser state dominate conversions and invariant accesses"
)]
// Codec declarations and string emitters follow schema traversal order. These
// bounded serializers favor direct append operations and explicit cases so the
// output can be audited against the standard.
#![allow(
    clippy::allow_attributes_without_reason,
    clippy::arbitrary_source_item_ordering,
    clippy::doc_markdown,
    clippy::format_push_string,
    clippy::if_same_then_else,
    clippy::items_after_statements,
    clippy::match_same_arms,
    clippy::unnecessary_wraps,
    clippy::unreadable_literal,
    reason = "DrawingML codecs retain schema order and bounded direct XML emission for specification auditability"
)]

pub mod blip;
pub mod chart;
pub mod color;
pub mod coordinate;
pub mod diagram;
pub mod ext;
pub mod fill;
pub mod geom;
pub mod geometry;
pub mod model3d;
pub mod text;
pub mod theme;
pub mod transform;

use thiserror::Error;

/// Result of a `DrawingML` read operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Failure to parse a shared `DrawingML` primitive.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The XML is malformed or contains an invalid encoded value.
    #[error("invalid DrawingML XML: {0}")]
    Xml(String),

    /// The document violates `DrawingML` structural or value constraints.
    #[error("invalid DrawingML: {0}")]
    Invalid(String),

    /// Reading a caller-provided chart stream failed.
    #[error("DrawingML input error: {0}")]
    Io(#[from] std::io::Error),

    /// Markup-compatibility preprocessing failed.
    #[error("DrawingML markup compatibility error: {0}")]
    Mce(#[from] litchi_ooxml_common::mce::Error),

    /// Shared OOXML entity or namespace decoding failed.
    #[error("DrawingML shared XML error: {0}")]
    Common(#[from] litchi_ooxml_common::XmlError),

    /// A bounded `DrawingML` codec resource was exhausted.
    #[error("DrawingML {resource} exceeds the limit of {limit}")]
    Limit {
        /// Resource whose configured bound was exceeded.
        resource: &'static str,
        /// Active upper bound.
        limit: usize,
    },
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}
