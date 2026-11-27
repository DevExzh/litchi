//! Explicit physical metadata of a parsed shape.
//!
//! Native relationship IDs and XML byte spans are package-local wire data,
//! not ordinary shape identities or unified node properties.
pub use super::model::{Payload, PayloadAttribute};
use super::{Shape, Span};

/// Relationship through which this shape references its payload part.
///
/// # Errors
/// Refuses competing payloads and invalid relationship attribute namespaces.
pub fn payload(shape: Shape<'_>) -> Result<Option<Payload<'_>>, Error> {
    shape.common().payload()
}

/// Span of the shape's non-visual property start tag in processed owner XML.
#[must_use]
pub fn non_visual_span(shape: Shape<'_>) -> Option<Span> {
    shape.common().non_visual_span()
}

/// Span of the shape's transform in processed owner XML.
#[must_use]
pub fn transform_span(shape: Shape<'_>) -> Option<Span> {
    shape.common().transform_span()
}

/// Span of the shape's direct `p:txBody` in processed owner XML.
///
/// This exposes source coordinates only; it does not decode text or apply
/// namespace and duplicate-body validation. Use the ordinary shape body
/// reader for semantic text.
#[must_use]
pub fn text_body_span(shape: Shape<'_>) -> Option<Span> {
    shape.common().text_body_span()
}

/// Content-free failures reading native payload relationship metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// Multiple competing relationship attributes or payload elements.
    #[error("shape payload relationship is ambiguous")]
    Ambiguous,
    /// Malformed attribute syntax or value.
    #[error("shape payload relationship attribute is malformed")]
    InvalidAttribute,
    /// A relationship prefix is undeclared or uses the wrong document dialect.
    #[error("shape payload relationship namespace is invalid")]
    InvalidNamespace,
}
