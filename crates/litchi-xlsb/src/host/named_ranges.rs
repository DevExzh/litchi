//! XLSB defined-name boundary.
//!
//! `litchi-xlsb` owns the semantic [`Definition`] model and bounded `BrtName`
//! codec. The host only exposes those contextual APIs and maps owner errors
//! into its package-level error type.

use crate::named_ranges as owner;
use crate::package::error::Error as HostError;

pub use owner::{
    Definition, Error, MAX_RECORD_BYTES, Result, area3d_formula, parse, validate_name,
};

impl From<Error> for HostError {
    fn from(error: Error) -> Self {
        match error {
            Error::Wire(error) => Self::from(error),
            Error::Formula(error) => Self::from(error),
            Error::InvalidLength { expected, found } => Self::InvalidLength { expected, found },
            Error::InvalidFormula(message) => Self::InvalidFormula(message),
        }
    }
}
