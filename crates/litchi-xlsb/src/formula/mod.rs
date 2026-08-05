//! Layered XLSB formula owner.
//!
//! `model` contains the semantic Ptg and formula data types. `codec` owns
//! BIFF12 parsing, serialization, and text compilation while leaving workbook
//! relationship resolution to the host adapter.

use thiserror::Error;

mod codec;
mod function_table;
mod model;

/// Maximum size of an XLSB cell formula token stream.
///
/// [MS-XLSB] 2.5.98.4 requires `cce` to be less than 16,385 bytes.
pub const MAX_CELL_FORMULA_BYTES: usize = 16_384;

/// Error returned by the standalone formula codec.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// A formula or token violates the BIFF12 formula grammar.
    #[error("invalid formula: {0}")]
    InvalidFormula(String),
    /// A cell or range coordinate is outside the Excel grid.
    #[error("invalid cell reference: {0}")]
    InvalidCellReference(String),
    /// A fixed-width payload is shorter than the required structure.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    /// A formula feature is valid but not supported by this codec.
    #[error("unsupported formula feature: {0}")]
    UnsupportedFeature(String),
    /// Text or primitive binary decoding failed.
    #[error("formula encoding: {0}")]
    Encoding(String),
}

/// Result type for standalone formula codecs.
pub type Result<T> = std::result::Result<T, Error>;

pub use codec::{Compiler, Parser, Resolution};
pub use model::*;
