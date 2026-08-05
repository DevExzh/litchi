//! Checked spreadsheet coordinates and A1 range notation.
//!
//! The semantic model is kept independent from its textual A1 codec.  Public
//! callers use [`CellCoord`] and [`CellRange`] for validated values; the
//! conversion helpers remain available for small format adapters.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use codec::{alpha_to_digit, digit_to_alpha};
pub use model::{CellCoord, CellRange, MAX_INDEX};
