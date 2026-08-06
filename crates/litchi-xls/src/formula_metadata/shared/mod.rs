//! Checked BIFF8 shared-formula ownership.
//!
//! The owner is deliberately inert: it stores the shared parsed formula and
//! the cells that use it, but never evaluates or expands the expression.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{Record, parse};
pub use model::{Cell, Owner, Range};

pub(crate) use validation::{FIXED_PAYLOAD_SIZE, MAX_FORMULA_BYTES};
