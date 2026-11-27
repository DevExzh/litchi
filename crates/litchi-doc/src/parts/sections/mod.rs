//! Layered owner for Word 97+ section ranges and section properties.
//!
//! The public facade exposes the typed section table while binary decoding,
//! wire validation, and regression coverage remain in their contextual layers.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::SectionsTable;
