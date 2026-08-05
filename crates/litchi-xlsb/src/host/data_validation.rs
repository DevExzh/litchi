//! OOXML-facing adapter for the canonical XLSB data-validation codec.
//!
//! `litchi_xlsb` owns the bounded `BrtDVal`/`BrtDVal14` record grammar and
//! formula-preserving semantic model. The OOXML host supplies its formula
//! representation, error surface, and worksheet/package orchestration.

use crate::data_validation as owner;
use crate::formula::ParsedFormula;
use crate::package::error::{Error, Result};

pub use owner::{RecordKind, Settings};

/// Validation rule bound to the OOXML host formula representation.
pub type Validation = owner::Validation<ParsedFormula>;

impl From<owner::Error> for Error {
    fn from(error: owner::Error) -> Self {
        match error {
            owner::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            owner::Error::InvalidFormula(message) => Self::InvalidFormula(message),
            owner::Error::InvalidCellReference(reference) => Self::InvalidCellReference(reference),
            owner::Error::Encoding(message) => Self::Encoding(message),
            owner::Error::UnsupportedFeature(feature) => Self::UnsupportedFeature(feature),
            owner::Error::Unrecognized { typ, val } => Self::Unrecognized { typ, val },
            owner::Error::Wire(error) => Self::Wire(error),
            owner::Error::Formula(error) => Self::from(error),
            owner::Error::Io(error) => Self::Io(error),
        }
    }
}

pub(crate) fn parse_collection_settings(data: &[u8], extension14: bool) -> Result<(Settings, u32)> {
    owner::parse_collection_settings(data, extension14).map_err(Into::into)
}

pub(crate) fn parse_dval_list(data: &[u8]) -> Result<String> {
    owner::parse_dval_list(data).map_err(Into::into)
}

pub(crate) fn validate_dval_list_formula(value: &str) -> Result<()> {
    owner::validate_dval_list_formula(value).map_err(Into::into)
}
