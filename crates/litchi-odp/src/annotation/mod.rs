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

pub use litchi_odf_common::annotation::Annotation;
pub use model::{Anchor, Info, Position};

pub(crate) use codec::{add, annotations, find, remove, replace};
