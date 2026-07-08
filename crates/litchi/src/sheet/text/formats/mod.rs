//! Format-specific implementations for text workbooks.

pub mod delimited;
pub mod dif;
pub mod fixed_width;
pub mod sylk;

// Re-export common types and functions
pub use delimited::{DelimitedConfig, read_delimited, write_delimited};
pub use dif::{DifConfig, read_dif, write_dif};
pub use fixed_width::{FixedWidthConfig, read_fixed_width, write_fixed_width};
pub use sylk::{SylkConfig, read_sylk, write_sylk};

#[cfg(test)]
mod tests;
