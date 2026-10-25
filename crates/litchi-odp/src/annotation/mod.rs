//! Slide- and shape-anchored ODF annotations.
//!
//! The rich annotation body is owned by `litchi-odf-common`; this module owns
//! only the presentation-specific anchor vocabulary and the lossless XML
//! editing seam.  Package identifiers and XML spans remain private.

mod codec;
mod model;
mod package;
mod validation;

#[cfg(test)]
mod tests;

use litchi_odf_common::XmlStreamLimits;
pub use litchi_odf_common::annotation::Annotation;
pub use model::{Anchor, Info, Position};

pub(crate) use codec::{add, annotations, find, remove, replace};

/// Inventory the annotations of a presentation `content.xml` (or a flat
/// presentation document) in document order, without opening a package.
///
/// # Errors
///
/// Returns an error when the XML is malformed or an annotation anchor is
/// invalid.
pub fn inventory(content_xml: &str, limits: &XmlStreamLimits) -> litchi_core::Result<Vec<Info>> {
    annotations(content_xml, limits)
}
