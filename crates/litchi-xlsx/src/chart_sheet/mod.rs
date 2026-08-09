//! Typed `SpreadsheetML` chartsheet APIs.
//!
//! The semantic model and bounded XML codec are layered separately from the
//! OPC package graph. The historical `litchi_xlsx::chart_sheet::*` facade is
//! retained through the re-exports below.

mod codec;
mod model;

/// Chartsheet package/resource graph and bounded load/store operations.
pub mod package;

#[cfg(test)]
mod tests;

pub use codec::{parse_chartsheet, validate_chartsheet, write_chartsheet};
pub use model::*;
