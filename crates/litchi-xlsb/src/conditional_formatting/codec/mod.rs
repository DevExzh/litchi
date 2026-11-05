//! Layered XLSB conditional-formatting codec facade.
//!
//! The semantic validator and bounded Brt* binary codec are kept in
//! contextual owners while this module preserves the public facade.

#![allow(
    clippy::too_many_arguments,
    reason = "arguments mirror independent BIFF12 conditional-formatting fields"
)]

mod binary;
mod semantic;
#[cfg(test)]
mod tests;

use thiserror::Error;

/// Result type for conditional-formatting codecs.
pub type Result<T> = std::result::Result<T, Error>;

/// Strict conditional-formatting codec error.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    #[error("invalid formula: {0}")]
    InvalidFormula(&'static str),
    #[error("invalid cell reference: {0}")]
    InvalidCellReference(&'static str),
    #[error("encoding error: {0}")]
    Encoding(&'static str),
    #[error("unsupported feature: {0}")]
    UnsupportedFeature(&'static str),
    #[error("unrecognized {typ}: {val}")]
    Unrecognized {
        typ: &'static str,
        val: &'static str,
    },
    #[error(transparent)]
    Wire(#[from] crate::raw::Error),
    #[error(transparent)]
    Formula(#[from] crate::formula::Error),
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}

pub(super) fn invalid(typ: &'static str, val: &'static str) -> Error {
    Error::Unrecognized { typ, val }
}

pub use binary::{
    parse_classic_header, parse_rule_extension_guid, serialize_rule_extension_guid,
    write_conditional_formattings,
};
pub use semantic::{icon_count14, validate_formula_count, validate_template};
