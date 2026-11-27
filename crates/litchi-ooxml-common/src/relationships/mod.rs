//! Typed OOXML relationship-namespace vocabulary.
//!
//! The semantic [`Id`] model is kept separate from namespace-aware XML
//! decoding so package owners can use checked identifiers without coupling
//! their APIs to a particular XML reader.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::{STRICT_NAMESPACE, TRANSITIONAL_NAMESPACE, attribute_id, attribute_value};
pub use model::Id;
