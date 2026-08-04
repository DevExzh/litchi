//! Compatibility path for typed XLSB conditional-formatting records.
//!
//! The strict semantic models and binary codecs are owned by the concrete
//! XLSB crate. This adapter preserves the historical host module path and
//! translates owner errors at the graph boundary; worksheet/package
//! orchestration remains in litchi-ooxml.

use crate::xlsb::error::XlsbError as HostXlsbError;

pub use litchi_xlsb::conditional_formatting::*;

impl From<litchi_xlsb::conditional_formatting::Error> for HostXlsbError {
    fn from(error: litchi_xlsb::conditional_formatting::Error) -> Self {
        match error {
            litchi_xlsb::conditional_formatting::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            litchi_xlsb::conditional_formatting::Error::InvalidFormula(message) => {
                Self::InvalidFormula(message)
            },
            litchi_xlsb::conditional_formatting::Error::InvalidCellReference(reference) => {
                Self::InvalidCellReference(reference)
            },
            litchi_xlsb::conditional_formatting::Error::Encoding(message) => {
                Self::Encoding(message)
            },
            litchi_xlsb::conditional_formatting::Error::UnsupportedFeature(feature) => {
                Self::UnsupportedFeature(feature)
            },
            litchi_xlsb::conditional_formatting::Error::Unrecognized { typ, val } => {
                Self::Unrecognized { typ, val }
            },
            litchi_xlsb::conditional_formatting::Error::Wire(error) => Self::from(error),
            litchi_xlsb::conditional_formatting::Error::Formula(error) => Self::from(error),
            litchi_xlsb::conditional_formatting::Error::Io(error) => Self::Io(error),
            error => Self::Unrecognized {
                typ: "conditional-formatting".to_string(),
                val: error.to_string(),
            },
        }
    }
}
