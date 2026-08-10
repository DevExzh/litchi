//! Typed XLSX failures.

use std::collections::TryReserveError;

use thiserror::Error;

use litchi_sheet::{Cell as Address, Column as ColumnIndex, Rect, Row as RowIndex};

use crate::column::WidthError;
use crate::layout::{
    DescentError, HeightError as DefaultHeightError, WidthError as DefaultWidthError,
};
use crate::outline::OutlineError;
use crate::row::HeightError;
use crate::sheet::NameError;
use crate::workbook::JoinError;

/// Result of an XLSX operation.
pub type Result<T> = std::result::Result<T, Error>;

#[cold]
#[track_caller]
pub(crate) fn panic_missing_invariant(message: &str) -> ! {
    panic!("XLSX internal invariant failed: {message}")
}

#[cold]
#[track_caller]
pub(crate) fn panic_error_invariant(message: &str, error: impl std::fmt::Display) -> ! {
    panic!("XLSX internal invariant failed: {message}: {error}")
}

/// Failure to open, inspect, or edit an XLSX document.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// Runtime-neutral OOXML managed-package encryption failed.
    #[cfg(feature = "encryption")]
    #[error("OOXML encryption error: {0}")]
    Encryption(#[from] litchi_crypto::ooxml::Error),
    /// An operation would silently discard or incorrectly assume encryption
    /// provenance.
    #[cfg(feature = "encryption")]
    #[error("XLSX encryption policy rejected {operation}: {source}")]
    EncryptionPolicy {
        /// Operation rejected at the XLSX boundary.
        operation: &'static str,
        /// Host-independent provenance policy failure.
        #[source]
        source: litchi_ooxml_common::package_encryption::PolicyError,
    },
    /// The OPC package is malformed or inaccessible.
    #[error("OPC package error: {0}")]
    Package(#[from] litchi_opc::OpcError),
    /// Markup-compatibility preprocessing failed.
    #[error("markup compatibility error: {0}")]
    MarkupCompatibility(#[from] litchi_ooxml_common::mce::Error),
    /// Shared OOXML decoding failed.
    #[error("OOXML decoding error: {0}")]
    Xml(#[from] litchi_ooxml_common::XmlError),
    /// `DrawingML` chart or geometry decoding failed.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::Error),
    /// A host-neutral OOXML package service failed.
    #[error("shared OOXML service error: {0}")]
    Common(#[from] litchi_ooxml_common::Error),
    /// A spreadsheet coordinate lies outside its typed domain.
    #[error(transparent)]
    Coordinate(#[from] litchi_sheet::CoordinateError),
    /// A spreadsheet range is empty, inverted, or outside its typed domain.
    #[error(transparent)]
    Range(#[from] litchi_sheet::RangeError),
    /// A column width is non-finite or outside Office's checked domain.
    #[error(transparent)]
    ColumnWidth(#[from] WidthError),
    /// A row height is non-finite or outside Excel's checked domain.
    #[error(transparent)]
    RowHeight(#[from] HeightError),
    /// A worksheet default row height is negative or non-finite.
    #[error(transparent)]
    DefaultHeight(#[from] DefaultHeightError),
    /// A worksheet default column width is outside Office's checked domain.
    #[error(transparent)]
    DefaultWidth(#[from] DefaultWidthError),
    /// A typographic descent is negative or non-finite.
    #[error(transparent)]
    Descent(#[from] DescentError),
    /// A row or column outline level is outside Office's checked domain.
    #[error(transparent)]
    Outline(#[from] OutlineError),
    /// A worksheet name is outside Office's checked domain.
    #[error(transparent)]
    SheetName(#[from] NameError),
    /// An XLSX structural invariant is invalid.
    #[error("invalid XLSX structure: {0}")]
    Invalid(String),
    /// A bounded XLSX operation could not reserve its required memory.
    #[error("could not reserve memory for XLSX {resource}: {source}")]
    Allocation {
        /// Resource whose bounded plan could not be reserved.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
    },
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
    /// A worksheet Office Add-in binding has no unique package-level target.
    #[error(
        "worksheet '{sheet}' Office Add-in binding '{app_ref}' has no package-level MS-OWEXML binding"
    )]
    DanglingWebBinding { sheet: String, app_ref: String },
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
    /// A safe worksheet-default edit is forbidden by worksheet state.
    #[error("cannot edit defaults on sheet '{sheet}': {reason}")]
    DefaultsEditBlocked {
        sheet: String,
        reason: DefaultsEditBlock,
    },
    /// A safe merged-range edit would lose content or cross unmodeled state.
    #[error("cannot edit merged range {range} on sheet '{sheet}': {reason}")]
    MergeEditBlocked {
        sheet: String,
        range: Rect,
        reason: MergeEditBlock,
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
    /// A bounded durable patch envelope was invalid or exceeded its limits.
    #[error("durable XLSX patch error: {0}")]
    DurablePatch(#[from] litchi_core::PatchError),
    /// Independently prepared edits could not be joined.
    #[error(transparent)]
    Join(Box<JoinError>),
    /// A selector variant is not supported by this API version.
    #[error("unsupported sheet selector")]
    UnsupportedSelector,
    /// The package owner intentionally does not execute host/runtime
    /// behavior such as recalculation, refresh, or rendering.
    #[error("unsupported XLSX operation: {feature}")]
    Unsupported { feature: &'static str },
}

impl From<std::convert::Infallible> for Error {
    fn from(never: std::convert::Infallible) -> Self {
        match never {}
    }
}

impl From<litchi_spreadsheet_drawing::Error> for Error {
    fn from(error: litchi_spreadsheet_drawing::Error) -> Self {
        match error {
            litchi_spreadsheet_drawing::Error::Drawing(error) => Self::Drawing(error),
            litchi_spreadsheet_drawing::Error::Invalid(message)
            | litchi_spreadsheet_drawing::Error::Encoding(message) => Self::Invalid(message),
            other @ (litchi_spreadsheet_drawing::Error::Mce(_)
            | litchi_spreadsheet_drawing::Error::Xml(_)
            | _) => Self::Invalid(other.to_string()),
        }
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
    /// A new column record cannot safely carry a style without an explicit
    /// width because Excel interprets that style-only record as zero-width.
    StyleNeedsWidth,
}

/// Why the ordinary worksheet-default editor refused a mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DefaultsEditBlock {
    ProtectedSheet,
    MarkupCompatibility,
    /// Materializing `sheetFormatPr` requires its mandatory row height.
    NeedsHeight,
}

/// Why the ordinary merged-range editor refused a mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum MergeEditBlock {
    ProtectedSheet,
    SingleCell,
    Overlap { existing: Rect },
    FollowerContent { address: Address },
    GroupFormula,
    MarkupCompatibility,
    UnmodeledPayload,
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
            Self::StyleNeedsWidth => {
                "styling an implicit column requires an explicit width in the same transaction"
            },
        })
    }
}

impl std::fmt::Display for DefaultsEditBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::ProtectedSheet => "the worksheet is protected",
            Self::MarkupCompatibility => {
                "the effective defaults are controlled by unmodeled compatibility markup"
            },
            Self::NeedsHeight => {
                "creating worksheet defaults requires a height in the same transaction"
            },
        })
    }
}

impl std::fmt::Display for MergeEditBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ProtectedSheet => formatter.write_str("the worksheet is protected"),
            Self::SingleCell => formatter.write_str("a merge must contain at least two cells"),
            Self::Overlap { existing } => {
                write!(formatter, "it overlaps existing merged range {existing}")
            },
            Self::FollowerContent { address } => write!(
                formatter,
                "covered cell {address} has content that Excel could discard"
            ),
            Self::GroupFormula => {
                formatter.write_str("it intersects an array, data-table, or shared formula group")
            },
            Self::MarkupCompatibility => formatter.write_str(
                "effective merged ranges are controlled by unmodeled compatibility markup",
            ),
            Self::UnmodeledPayload => {
                formatter.write_str("the merge container has unmodeled child payload")
            },
        }
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

pub(crate) fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use std::error::Error as _;

    use super::{Error, allocation};

    #[test]
    fn allocation_preserves_resource_and_source() {
        let mut values = Vec::<u8>::new();
        let source = values
            .try_reserve(usize::MAX)
            .expect_err("usize::MAX must exceed Vec's maximum capacity");
        let error = allocation("test buffer", source);

        assert!(matches!(
            &error,
            Error::Allocation { resource, .. } if *resource == "test buffer"
        ));
        assert!(error.source().is_some());
        assert!(error.to_string().contains("XLSX test buffer"));
    }
}
