//! Typed, inert `SpreadsheetML` calculation-chain metadata and package ownership.
//!
//! Semantic values live in `model`, XML grammar in `codec`, and OPC graph
//! ownership in `package`. The facade keeps the ADR 0018 API compact.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{read, write};
pub use model::{Cell, Chain, Conformance, Flags, Sheet, Step, raw};
pub use package::{load, put, remove};

pub(crate) use package::validate_package;
