//! Compatibility path for the typed, inert XLSB External Link codec.
//!
//! The semantic model and bounded `BrtSupBook` stream parser now live in
//! [`litchi_xlsb::external_link`]. OPC relationship validation and workbook
//! graph orchestration remain in this host crate.

pub use litchi_xlsb::external_link::*;

impl From<litchi_xlsb::external_link::Error> for crate::xlsb::error::Error {
    fn from(error: litchi_xlsb::external_link::Error) -> Self {
        match error {
            litchi_xlsb::external_link::Error::Wire(error) => Self::Wire(error),
            litchi_xlsb::external_link::Error::InvalidFormula(message) => {
                Self::InvalidFormula(message)
            },
            litchi_xlsb::external_link::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            litchi_xlsb::external_link::Error::Allocation { resource, source } => {
                Self::Allocation { resource, source }
            },
            error => Self::InvalidFormula(error.to_string()),
        }
    }
}
