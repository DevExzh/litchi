//! Typed XLSX failures.

use thiserror::Error;

use litchi_sheet::{Cell as Address, Column as ColumnIndex, Row as RowIndex};

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
    /// A spreadsheet range is empty, inverted, or outside its typed domain.
    #[error(transparent)]
    Range(#[from] litchi_sheet::RangeError),
    /// An XLSX structural invariant is invalid.
    #[error("invalid XLSX structure: {0}")]
    Invalid(String),
    /// A requested name matches more than one sheet.
    #[error("sheet name '{name}' is ambiguous ({matches} matches)")]
    AmbiguousSheetName { name: String, matches: usize },
    /// A shared-style handle belongs to an unrelated resource-table lineage.
    #[error("shared style belongs to a different resource-table lineage")]
    ForeignStyle,
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
    /// A safe row-property edit is forbidden by worksheet state.
    #[error("cannot edit row index {row} on sheet '{sheet}': {reason}")]
    RowEditBlocked {
        sheet: String,
        row: RowIndex,
        reason: RowEditBlock,
    },
    /// A safe column-property edit is forbidden by worksheet state or by an
    /// extension payload that cannot be split without changing its meaning.
    #[error("cannot edit column index {column} on sheet '{sheet}': {reason}")]
    ColumnEditBlocked {
        sheet: String,
        column: ColumnIndex,
        reason: ColumnEditBlock,
    },
    /// A safe workbook tab edit is forbidden by workbook structure or by an
    /// extension payload whose effective catalog cannot be edited losslessly.
    #[error("cannot edit tab at position {position} ('{sheet}'): {reason}")]
    TabEditBlocked {
        sheet: String,
        position: usize,
        reason: TabEditBlock,
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

/// Why the ordinary row editor refused a property mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RowEditBlock {
    ProtectedSheet,
}

/// Why the ordinary column editor refused a property mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ColumnEditBlock {
    ProtectedSheet,
    MarkupCompatibility,
}

/// Why the workbook tab editor refused a visibility mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum TabEditBlock {
    LastVisibleTab,
    ProtectedWorkbook,
    MarkupCompatibility,
}

impl std::fmt::Display for TabEditBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::LastVisibleTab => "the workbook must retain at least one visible tab",
            Self::ProtectedWorkbook => "the workbook structure is protected",
            Self::MarkupCompatibility => {
                "the effective workbook catalog contains unmodeled compatibility markup"
            },
        })
    }
}

impl std::fmt::Display for ColumnEditBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::ProtectedSheet => "the worksheet is protected",
            Self::MarkupCompatibility => {
                "the effective column record contains an unmodeled extension payload"
            },
        })
    }
}

impl std::fmt::Display for RowEditBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::ProtectedSheet => "the worksheet is protected",
        })
    }
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
