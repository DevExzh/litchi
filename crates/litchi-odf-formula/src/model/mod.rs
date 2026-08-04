//! In-memory MathML model.

pub use crate::migration::document::{Attribute, Content, Element, Kind};

/// The namespace used by MathML presentation and content elements.
pub const NAMESPACE: &str = crate::migration::document::MATHML_NAMESPACE;
