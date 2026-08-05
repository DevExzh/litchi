//! Compatibility path for the typed, inert XLSB External Link codec.
//!
//! The semantic model and bounded `BrtSupBook` stream parser now live in
//! [`crate::external_link`]. OPC relationship validation and workbook
//! graph orchestration remain in this host crate.

use crate::external_link as owner;
use crate::package::error::Error as HostError;

pub use owner::*;

impl From<Error> for HostError {
    fn from(error: Error) -> Self {
        match error {
            Error::Wire(error) => Self::Wire(error),
            Error::InvalidFormula(message) => Self::InvalidFormula(message),
            Error::InvalidLength { expected, found } => Self::InvalidLength { expected, found },
            Error::Allocation { resource, source } => Self::Allocation { resource, source },
        }
    }
}
