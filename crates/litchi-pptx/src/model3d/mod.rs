//! Contextual PresentationML 3D-model ownership.
//!
//! The shared [`litchi_drawingml::model3d`] grammar owns the `am3d:model3d`
//! sequence and its bounded inert children.  This module adds the PPTX
//! graphic-frame anchor, OPC relationship policy, opaque GLB/image resources,
//! and atomic slide-level CRUD without exposing relationship IDs or part
//! names through the ordinary facade.

pub mod codec;
pub mod model;
pub(crate) mod package;
pub mod validation;

#[cfg(test)]
mod tests;

pub use model::{Asset, Data, Link, Model, Preview, Scene, Shape, Unknown};

/// The `a:graphicData/@uri` value used by a PresentationML 3D model frame.
pub const GRAPHIC_URI: &str = litchi_drawingml::model3d::NAMESPACE;

/// Relationship type used by Office for a model3d binary asset.
pub(crate) const MODEL_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/model3d";
/// Content type used by Office for an embedded glTF binary model.
pub(crate) const MODEL_CONTENT_TYPE: &str = "model/gltf-binary";

pub(crate) const MAX_SLIDE_XML_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_MODELS_PER_SLIDE: usize = 4_096;
pub(crate) const MAX_SHAPES_PER_SLIDE: usize = 100_000;
pub(crate) const MAX_MODEL_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_PREVIEW_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_TOTAL_PAYLOAD_BYTES: usize = 96 * 1024 * 1024;
pub(crate) const MAX_LINK_BYTES: usize = 16 * 1024;
pub(crate) const MAX_XML_DEPTH: usize = 256;
