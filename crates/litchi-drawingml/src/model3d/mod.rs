//! Host-neutral metadata for a DrawingML 3D-model instance.
//!
//! `[MS-ODRAWXML]` §2.31 stores the model payload, camera, transform, and
//! viewport as a `model3d` element.  This first shared slice gives callers a
//! typed view of the two relationship-bearing `AG_Blob` attributes and the
//! raster preview while retaining every other child as bounded inert XML.
//! The 3D scene itself remains data, not a renderer or an activation path.
//!
//! The semantic model, XML wire codec, package relationship seam, validation,
//! and conformance tests are deliberately separate.  Concrete DOCX, PPTX,
//! XLSX, and XLSB crates supply their own package adapter through
//! [`Resolver`]; this crate never owns a host anchor or an OPC package.

pub mod codec;
mod model;
pub mod package;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    Blip, Child, Id, Inert, Metadata, Namespace, Raster, RasterChild, Reference, ValueError,
};
pub use package::{Relationship, Resolver, Target};
pub use validation::{Error, validate, validate_relationships};

/// Microsoft 2017 3D-model namespace defined by `[MS-ODRAWXML]` §2.31.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2017/model3d";
/// Transitional DrawingML namespace used by the imported `a:` vocabulary.
pub const DRAWING_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
/// Strict DrawingML namespace used by the imported `a:` vocabulary.
pub const DRAWING_NAMESPACE_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
/// Transitional OOXML relationship namespace.
pub const RELATIONSHIP_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// Strict OOXML relationship namespace.
pub const RELATIONSHIP_NAMESPACE_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";

pub(super) const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_FRAGMENT_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_CHILDREN: usize = 4_096;
pub(super) const MAX_RASTER_CHILDREN: usize = 512;
pub(super) const MAX_DEPTH: usize = 128;
pub(super) const MAX_NODES: usize = 100_000;
pub(super) const MAX_RELATIONSHIP_ID_BYTES: usize = 256;
pub(super) const MAX_RENDERER_TEXT_BYTES: usize = 4_096;
pub(super) const MAX_NAMESPACE_DECLARATIONS: usize = 256;
pub(super) const MAX_NAMESPACE_TEXT_BYTES: usize = 4_096;
