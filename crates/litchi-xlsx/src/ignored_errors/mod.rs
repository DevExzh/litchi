//! Immutable XLSX worksheet ignored-error read model.
//!
//! The owner is layered by responsibility: public semantic models, bounded XML
//! conversion, and focused regression coverage.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::parse_worksheet_ignored_errors;
pub use model::*;
