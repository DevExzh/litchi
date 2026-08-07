//! In-memory `MathML` model.

mod edit;
mod element;

pub(crate) use element::MATHML_NAMESPACE;
pub use element::{Attribute, Content, Element, Kind};

/// The namespace used by `MathML` presentation and content elements.
pub const NAMESPACE: &str = MATHML_NAMESPACE;
