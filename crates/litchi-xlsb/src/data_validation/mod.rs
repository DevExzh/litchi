//! Typed, bounded XLSB data-validation records.
//!
//! The layouts and invariants in this module follow [MS-XLSB] sections
//! 2.4.55–2.4.58, 2.5.36–2.5.37, 2.5.58–2.5.66, 2.5.98.8, and 2.5.156.
//! The semantic model is separated from the BIFF12 codec; worksheet/package
//! traversal stays in the OOXML host.

use std::io;
use thiserror::Error as ThisError;

mod codec;
mod model;

#[cfg(test)]
mod tests;

/// Result type for the owner data-validation codecs.
pub type Result<T> = std::result::Result<T, Error>;

/// Strict data-validation codec error.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A record field or payload is truncated or has trailing bytes.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    /// A formula stream violates the bounded BIFF12 formula contract.
    #[error("invalid formula: {0}")]
    InvalidFormula(String),
    /// A cell or target range is outside the worksheet grid.
    #[error("invalid cell reference: {0}")]
    InvalidCellReference(String),
    /// UTF-16 or other text encoding failed.
    #[error("encoding error: {0}")]
    Encoding(String),
    /// A text formula construct is valid XLSB but not supported by the owner
    /// fallback compiler; callers may use the host compiler bridge.
    #[error("unsupported feature: {0}")]
    UnsupportedFeature(String),
    /// A record invariant or reserved field is invalid.
    #[error("unrecognized {typ}: {val}")]
    Unrecognized { typ: String, val: String },
    /// A raw BIFF12 boundary check failed.
    #[error(transparent)]
    Wire(#[from] crate::raw::Error),
    /// Formula token parsing or rendering failed.
    #[error(transparent)]
    Formula(#[from] crate::formula::Error),
    /// Writer I/O failed.
    #[error("I/O error: {0}")]
    Io(#[from] io::Error),
}

pub use codec::{parse_collection_settings, parse_dval_list, validate_dval_list_formula};
pub use model::{FormulaBinary, RecordKind, Settings, Validation};
