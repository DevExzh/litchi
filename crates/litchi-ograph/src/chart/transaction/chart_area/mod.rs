//! Transactional edits for the fixed-size `[MS-OGRAPH]` `Chart` record.

pub(crate) mod codec;
mod model;
pub(crate) mod validation;

pub use model::Change;
pub(crate) use model::Request;

#[cfg(test)]
mod tests;
