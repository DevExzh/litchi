//! Transactional edits for the fixed-size `[MS-OGRAPH]` `ShtProps` record.
//!
//! The editor changes only the existing four-byte payload.  The empty
//! `PlotArea` record is part of the source identity but is never inserted,
//! removed, or reordered by this capability.

pub(crate) mod codec;
mod model;
pub(crate) mod validation;

pub use model::Change;
pub(crate) use model::Request;

#[cfg(test)]
mod tests;
