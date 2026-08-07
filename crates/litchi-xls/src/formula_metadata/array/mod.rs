//! Checked, inert ownership of BIFF8 `Array` formula records.
//!
//! Formula tokens and ancillary bytes are retained as data. This module never
//! evaluates formulas, resolves external references, or invokes functions.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use super::shared::{Cell, Range};
pub use model::{Cells, Limits, Owner};

pub(crate) use codec::parse_payload;
