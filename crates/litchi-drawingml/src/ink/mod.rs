//! Source-backed InkML and DrawingML ink extension support.
//!
//! The model is intentionally passive: it validates the XML structure and
//! exposes context, trace, brush, and action metadata, but never recognizes,
//! replays, renders, or executes handwriting. Source-backed documents retain
//! one immutable source allocation and expose spans into that source, while
//! metadata-only reads retain typed values and offsets without the XML input.

mod authoring;
mod codec;
mod model;
mod validation;

pub mod actions;

#[cfg(test)]
mod tests;

pub use authoring::{
    AuthoringLimits, BrushDraft, BrushPropertyDraft, ContextDraft, Draft, Prepared, TraceDraft,
};
pub use codec::{
    read, read_metadata, read_metadata_with_source_spans, read_shared,
    read_shared_with_source_spans, write, write_to,
};
pub use model::{
    BrushProperty, BrushPropertyName, Context, ContextKind, Document, EffectiveBrushProperty, Guid,
    InkEffect, Metadata, SemanticType, SourceSpan, Trace, ValueError,
};

/// InkML namespace used by the content part and ink actions.
pub const INKML_NAMESPACE: &str = "http://www.w3.org/2003/InkML";
/// Microsoft ink interpretation namespace.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/ink/2010/main";
/// DrawingML 2016 ink brush extension namespace.
pub const DRAWING_2016_NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2016/ink";
/// PowerPoint 2014 ink actions namespace.
pub const ACTION_NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2014/inkAction";
/// XML namespace used by xml:id.
pub const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";

/// Maximum complete InkML source bytes retained by one immutable document.
pub const MAX_SOURCE_BYTES: usize = 16 * 1024 * 1024;
/// Maximum scalar attribute value decoded by the InkML codecs.
pub const MAX_ATTRIBUTE_VALUE_BYTES: usize = 1_048_576;
/// Maximum canonical GUID lexical length.
pub const MAX_GUID_BYTES: usize = 38;
/// Maximum future/custom token length retained by a typed InkML scalar.
pub const MAX_TOKEN_BYTES: usize = 256;
/// Maximum nesting depth accepted by the parser.
pub const MAX_DEPTH: usize = 128;
/// Maximum element nodes accepted by the parser.
pub const MAX_NODES: usize = 100_000;
/// Maximum context metadata records retained.
pub const MAX_CONTEXTS: usize = 32_768;
/// Maximum trace locators retained.
pub const MAX_TRACES: usize = 65_536;
/// Maximum brush property records retained.
pub const MAX_BRUSH_PROPERTIES: usize = 65_536;

mod xml_characters;
