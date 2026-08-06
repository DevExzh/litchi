//! Transactional edits for the fixed-size `[MS-OGRAPH]` `Chart` record.

pub(crate) mod codec;
mod model;
pub(crate) mod validation;

/// Transactional edits for the fixed-width `[MS-OGRAPH]` `Series` metadata.
pub mod series_metadata;

pub use model::Change;
pub(crate) use model::Request;

#[cfg(test)]
mod tests;
