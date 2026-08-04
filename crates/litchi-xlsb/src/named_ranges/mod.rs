//! Layered XLSB defined-name values and `BrtName` codec.
//!
//! The semantic [`Definition`] model and bounded binary codec live in this
//! crate. Workbook ordering, sheet-scope resolution, and package traversal
//! remain responsibilities of the OOXML host adapter.

use crate::raw::Error as WireError;
use thiserror::Error as ThisError;

mod codec;
mod model;
#[cfg(test)]
mod tests;

/// Result type for the standalone defined-name owner.
pub type Result<T> = std::result::Result<T, Error>;

/// Error returned by the bounded defined-name codec.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A BIFF12 header, payload, scalar, or string failed raw validation.
    #[error(transparent)]
    Wire(#[from] WireError),
    /// A formula binary structure failed validation in the shared formula owner.
    #[error(transparent)]
    Formula(#[from] crate::formula::Error),
    /// A fixed-width record or bounded owner input has the wrong size.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    /// A defined-name invariant is invalid.
    #[error("invalid formula: {0}")]
    InvalidFormula(String),
}

pub use codec::{MAX_RECORD_BYTES, area3d_formula, parse};
pub use model::{Definition, validate_name};

// Historical host vocabulary retained as compatibility aliases. New code
// should use the contextual owner names above.
pub use codec::area3d_formula as create_area3d_formula;
pub use model::Definition as NamedRange;
pub use model::validate_name as validate_defined_name;
