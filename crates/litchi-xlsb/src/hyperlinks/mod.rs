//! Typed XLSB hyperlink records.
//!
//! The semantic [`Hyperlink`] model is separate from the BIFF12 wire codec;
//! worksheet and relationship orchestration remains in the host layers.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{Error, Hyperlink, PREFIX_LEN, Result};
