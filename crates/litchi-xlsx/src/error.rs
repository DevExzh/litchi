//! Typed XLSX failures.

use thiserror::Error;

use litchi_sheet::{Cell as Address, Column as ColumnIndex, Row as RowIndex};

use crate::column::{OutlineError, WidthError};
use crate::sheet::NameError;
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
    /// A column width is non-finite or outside Office's checked domain.
    #[error(transparent)]
    ColumnWidth(#[from] WidthError),
    /// A column outline level is outside Office's checked domain.
    #[error(transparent)]
    ColumnOutline(#[from] OutlineError),
    /// A worksheet name is outside Office's checked domain.
    #[error(transparent)]
    SheetName(#[from] NameError),
    /// An XLSX structural invariant is invalid.
    #[error("invalid XLSX structure: {0}")]
    Invalid(String),
    /// Two logical sheets would have the same locale-independent name.
    #[error(
        "sheet name '{name}' conflicts between positions {first} and {second} under Unicode caseless matching"
    )]
    SheetNameConflict {
        name: String,
        first: usize,
        second: usize,
    },
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
    /// A sheet rename found a local reference that cannot be updated without
    /// guessing at an unmodeled producer extension.
    #[error(
        "cannot rename tab at position {position} ('{sheet}') while processing '{part}': {reason}"
    )]
    RenameBlocked {
        sheet: String,
        position: usize,
        part: String,
        reason: RenameBlock,
    },
    /// Removing a worksheet would discard a live dependency or cross a graph
    /// boundary whose semantics are not modeled by the safe facade.
    #[error(
        "cannot remove tab at position {position} ('{sheet}') while processing '{part}': {reason}"
    )]
    SheetRemoveBlocked {
        sheet: String,
        position: usize,
        part: String,
        reason: RemoveBlock,
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

impl From<std::convert::Infallible> for Error {
    fn from(never: std::convert::Infallible) -> Self {
        match never {}
    }
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

/// Why the workbook tab editor refused a state or ordering mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum TabEditBlock {
    SheetLimit,
    LastVisibleTab,
    NotVisible,
    ActiveTabLimit,
    ViewIndex,
    TrackedWorkbook,
    ProtectedWorkbook,
    MarkupCompatibility,
}

/// Why dependency-aware sheet rename refused one package part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RenameBlock {
    MarkupCompatibility,
    UnmodeledReference,
}

/// Why dependency-aware worksheet deletion refused to publish.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RemoveBlock {
    /// A workbook must retain at least one sheet.
    LastSheet,
    /// Worksheet removal is not yet composable with another mutation plan.
    MixedEdit,
    /// A retained, modeled formula or field still names the sheet.
    IncomingReference,
    /// An unknown producer field may still name the sheet.
    UnmodeledReference,
    /// A dependency occurs under markup-compatibility choice semantics.
    MarkupCompatibility,
    /// Macro code can address worksheets dynamically and is not rewritten.
    MacroProject,
    /// Another OPC relationship still targets the worksheet part.
    IncomingRelationship,
}

impl std::fmt::Display for RemoveBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::LastSheet => "a workbook must retain at least one sheet",
            Self::MixedEdit => {
                "worksheet removal cannot yet be combined with another mutation in one transaction"
            },
            Self::IncomingReference => "a retained modeled field still references the sheet",
            Self::UnmodeledReference => "an unmodeled producer field may still reference the sheet",
            Self::MarkupCompatibility => {
                "a reference is controlled by markup-compatibility content"
            },
            Self::MacroProject => "a VBA project may address the sheet dynamically",
            Self::IncomingRelationship => {
                "another package relationship still targets the worksheet part"
            },
        })
    }
}

impl std::fmt::Display for RenameBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::MarkupCompatibility => {
                "a local sheet reference is controlled by markup-compatibility content"
            },
            Self::UnmodeledReference => {
                "an unmodeled or inconsistent field can carry a local sheet reference"
            },
        })
    }
}

impl std::fmt::Display for TabEditBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::SheetLimit => "the workbook has reached Office's worksheet limit",
            Self::LastVisibleTab => "the workbook must retain at least one visible tab",
            Self::NotVisible => "only a visible tab can be active",
            Self::ActiveTabLimit => "the tab position exceeds Office's active-tab limit",
            Self::ViewIndex => "a workbook view contains an invalid sheet position",
            Self::TrackedWorkbook => "tracked workbooks require revision-aware tab reordering",
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
