//! Typed XLSX failures.

use thiserror::Error;

use litchi_sheet::Cell as Address;

use crate::workbook::JoinError;

/// Result of an XLSX operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Failure to open, inspect, or edit an XLSX document.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The OPC package is malformed or inaccessible.
    #[error("OPC package error: {0}")]
    Package(#[from] litchi_opc::OpcError),
    /// Markup-compatibility preprocessing failed.
    #[error("markup compatibility error: {0}")]
    MarkupCompatibility(#[from] litchi_ooxml_common::MceError),
    /// Shared OOXML decoding failed.
    #[error("OOXML decoding error: {0}")]
    Xml(#[from] litchi_ooxml_common::XmlError),
    /// A spreadsheet coordinate lies outside its typed domain.
    #[error(transparent)]
    Coordinate(#[from] litchi_sheet::CoordinateError),
    /// An XLSX structural invariant is invalid.
    #[error("invalid XLSX structure: {0}")]
    Invalid(String),
    /// A requested name matches more than one sheet.
    #[error("sheet name '{name}' is ambiguous ({matches} matches)")]
    AmbiguousSheetName { name: String, matches: usize },
    /// A worksheet-only operation targeted another sheet kind.
    #[error("sheet '{sheet}' is not a worksheet")]
    NotWorksheet { sheet: String },
    /// A safe edit requires a capability or dependency-aware operation that
    /// this transaction does not have.
    #[error("cannot edit {address} on sheet '{sheet}': {reason}")]
    EditBlocked {
        sheet: String,
        address: Address,
        reason: EditBlock,
    },
    /// Editing would invalidate an OPC digital signature.
    #[error("signed workbooks require explicit signature stripping before editing")]
    Signed,
    /// A patch expected different source bytes and was not applied.
    #[error("patch conflict in '{part}': target content differs from the expected state")]
    PatchConflict { part: String },
    /// Independently prepared edits could not be joined.
    #[error(transparent)]
    Join(Box<JoinError>),
    /// A selector variant is not supported by this API version.
    #[error("unsupported sheet selector")]
    UnsupportedSelector,
}

impl From<JoinError> for Error {
    fn from(error: JoinError) -> Self {
        Self::Join(Box::new(error))
    }
}

/// Why the ordinary cell editor refused a mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EditBlock {
    ProtectedSheet,
    CoveredMerge,
    GroupFormula,
    DataValidation,
    UnknownCell,
    MarkupCompatibility,
}

impl std::fmt::Display for EditBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::ProtectedSheet => "the worksheet is protected",
            Self::CoveredMerge => "the cell is covered by a merged range",
            Self::GroupFormula => "the cell belongs to a range-scoped formula",
            Self::DataValidation => "the cell has a validation rule that is not evaluated",
            Self::UnknownCell => "the existing cell encoding is not safely editable",
            Self::MarkupCompatibility => "the cell contains an unmodeled extension payload",
        })
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
