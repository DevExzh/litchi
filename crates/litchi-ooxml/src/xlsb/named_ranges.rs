//! XLSB defined-name boundary.
//!
//! `litchi-xlsb` owns the semantic [`Definition`] model and bounded `BrtName`
//! codec. The host only exposes those contextual APIs and maps owner errors
//! into its package-level error type.

pub use litchi_xlsb::named_ranges::{
    Definition, Error, MAX_RECORD_BYTES, Result, area3d_formula, parse, validate_name,
};

use crate::xlsb::error::XlsbError;

impl From<litchi_xlsb::named_ranges::Error> for XlsbError {
    fn from(error: litchi_xlsb::named_ranges::Error) -> Self {
        let message = error.to_string();
        match error {
            litchi_xlsb::named_ranges::Error::Wire(error) => Self::from(error),
            litchi_xlsb::named_ranges::Error::Formula(error) => Self::from(error),
            litchi_xlsb::named_ranges::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            litchi_xlsb::named_ranges::Error::InvalidFormula(message) => {
                Self::InvalidFormula(message)
            },
            _ => Self::InvalidFormula(message),
        }
    }
}
