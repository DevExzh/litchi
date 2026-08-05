//! Immutable workbook calculation-properties read model.
//!
//! The facade exposes the typed calculation policy while keeping its
//! SpreadsheetML/MCE representation and regression coverage in dedicated
//! layers.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::parse;
pub use model::{Mode, Properties, ReferenceMode};
