//! Typed metadata carried by a BIFF8 `Formula` record.
//!
//! The formula expression remains an opaque token stream owned by the cell
//! record. This owner types the surrounding calculation flags and application
//! cache without pretending to evaluate formulas or rebuild Excel's dependency
//! graph.

mod codec;
mod model;
pub mod shared;
mod validation;

#[cfg(test)]
mod tests;

pub use model::Metadata;
pub use shared::{Cell, Owner, Range};

pub(crate) use codec::parse_record;
pub(crate) use validation::{encode_flags, validate_for_write};
