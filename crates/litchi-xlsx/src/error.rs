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
    /// A shared sheet-name dependency scanner failed outside the legacy XLSX
    /// leaf mappings retained by the semantic transaction.
    #[error("sheet-name dependency error: {0}")]
    SheetNameDependency(#[source] litchi_ooxml_common::xml::sheet_name::Error),
    /// `DrawingML` chart or geometry decoding failed.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::Error),
    /// A host-neutral OOXML package service failed.
    #[error("shared OOXML service error: {0}")]
    Common(#[source] litchi_ooxml_common::Error),
    /// A caller-owned resource budget refused an operation.
    ///
    /// Retains the resource, observed amount, limit, and budget scope so callers
    /// can distinguish resource exhaustion from malformed document input.
    #[error(transparent)]
    ResourceLimit(#[from] litchi_core::ResourceLimit),
    /// A bound of an edit family was exceeded.
    #[error("XLSX edit limit exceeded for {kind}: {observed} exceeds {maximum}")]
    Limit {
        /// The bounded dimension and the family that owns it.
        kind: LimitKind,
        /// The observed amount.
        observed: u64,
        /// The maximum in force.
        maximum: u64,
    },
    /// A configured limit is outside the range its family accepts.
    #[error("invalid XLSX limit for {kind}: {value} exceeds {maximum}")]
    InvalidLimit {
        /// The configured dimension.
        kind: LimitKind,
        /// The configured value.
        value: u64,
        /// The largest accepted value.
        maximum: u64,
    },
    /// A staging or apply operation refused a well-formed request
    /// (record 0351 §2.4).
    #[error("XLSX {} refused ({}): {}", .0.operation(), .0.kind(), .0.reason())]
    Refused(litchi_core::edit::Refusal),
    /// The bounded form-control properties codec refused a payload or edit.
    #[error(transparent)]
    FormControl(#[from] crate::form_control::FormControlError),
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
    /// A semantic name selects multiple candidates under its identity policy.
    #[error("ambiguous {kind} selection: {matches} candidates")]
    Ambiguous {
        /// Structural kind of the selected object.
        kind: &'static str,
        /// Number of candidates satisfying the name.
        matches: usize,
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

    /// A source-backed package changed while the operation was in progress.
    #[error("XLSX source changed while the operation was in progress")]
    SourceChanged {
        /// Version captured before the operation.
        expected: litchi_core::SourceVersion,
        /// Version observed when the operation completed.
        observed: litchi_core::SourceVersion,
    },

    /// A generated staging or publication invariant was not satisfied.
    #[error("XLSX internal invariant failed: {0}")]
    Internal(&'static str),
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
    /// Removing this storage requires an explicit connection disposition.
    #[error("Custom Data storage '{id}' is referenced by {connections} connections")]
    CustomDataReferenced { id: String, connections: usize },
}

impl Error {
    /// A typed refusal of `operation` for `reason` (content-free text).
    pub(crate) const fn refused(
        kind: litchi_core::edit::RefusalKind,
        operation: &'static str,
        reason: &'static str,
    ) -> Self {
        Self::Refused(litchi_core::edit::Refusal::new(kind, operation, reason))
    }

    /// An exceeded bound of `kind`.
    pub(crate) fn limit(kind: impl Into<LimitKind>, observed: usize, maximum: usize) -> Self {
        Self::Limit {
            kind: kind.into(),
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        }
    }
}

impl From<litchi_ooxml_common::Error> for Error {
    fn from(error: litchi_ooxml_common::Error) -> Self {
        match error {
            litchi_ooxml_common::Error::Refused(refusal) => Self::Refused(refusal),
            litchi_ooxml_common::Error::PatchConflict => Self::PatchConflict {
                part: "shared OOXML graph".to_owned(),
            },
            litchi_ooxml_common::Error::Signed => Self::Signed,
            litchi_ooxml_common::Error::SourceChanged { expected, observed } => {
                Self::SourceChanged { expected, observed }
            },
            litchi_ooxml_common::Error::Internal(reason) => Self::Internal(reason),
            other => Self::Common(other),
        }
    }
}

/// A bounded dimension of an XLSX edit family, as reported by
/// [`Error::Limit`] and [`Error::InvalidLimit`].
///
/// Each variant wraps the typed `LimitKind` of the family that owns the
/// bound.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// The primary workbook transaction ([`crate::workbook::edit::LimitKind`]).
    Workbook(crate::workbook::edit::LimitKind),
    /// Conditional formatting ([`crate::conditional_formatting::edit::LimitKind`]).
    ConditionalFormatting(crate::conditional_formatting::edit::LimitKind),
    /// Data validation ([`crate::data_validation::edit::LimitKind`]).
    DataValidation(crate::data_validation::edit::LimitKind),
    /// Auto filters ([`crate::auto_filter::edit::LimitKind`]).
    AutoFilter(crate::auto_filter::edit::LimitKind),
    /// Legacy notes ([`crate::workbook::comments::edit::LimitKind`]).
    Comments(crate::workbook::comments::edit::LimitKind),
    /// Worksheet charts ([`crate::chart::edit::LimitKind`]).
    Chart(crate::chart::edit::LimitKind),
    /// Worksheet drawings ([`crate::drawing::edit::LimitKind`]).
    Drawing(crate::drawing::edit::LimitKind),
    /// Raw OPC edits ([`crate::opc::LimitKind`]).
    Opc(crate::opc::LimitKind),
    /// Custom Data ([`crate::custom_data::LimitKind`]).
    CustomData(crate::custom_data::LimitKind),
    /// Surveys ([`crate::survey::LimitKind`]).
    Survey(crate::survey::LimitKind),
    /// Workbook calculation metadata
    /// ([`crate::calculation_properties::LimitKind`]).
    CalculationProperties(crate::calculation_properties::LimitKind),
    /// Worksheet page breaks ([`crate::page_breaks::LimitKind`]).
    PageBreaks(crate::page_breaks::LimitKind),
    /// `pivotTableData` reads and edits
    /// ([`crate::pivot::table_data::LimitKind`]).
    PivotTableData(crate::pivot::table_data::LimitKind),
    /// PivotTable server formats ([`crate::pivot::server_formats::LimitKind`]).
    PivotServerFormats(crate::pivot::server_formats::LimitKind),
    /// Cached unique names ([`crate::pivot::cached_unique_names::LimitKind`]).
    CachedUniqueNames(crate::pivot::cached_unique_names::LimitKind),
    /// Custom XML Maps ([`crate::xml_maps::LimitKind`]).
    XmlMaps(crate::xml_maps::LimitKind),
    /// Source-backed cell values ([`crate::cell_values::LimitKind`]).
    CellValues(crate::cell_values::LimitKind),
    /// Source-backed row visibility
    /// ([`crate::row_visibility::source::LimitKind`]).
    RowVisibility(crate::row_visibility::source::LimitKind),
    /// Workbook revision metadata ([`crate::revisions::LimitKind`]).
    Revisions(crate::revisions::LimitKind),
    /// Data-type-icon XML and relationship budgets.
    DataTypeIcons(crate::data_type_icons::LimitKind),
    /// Workbook Data Model XML and payload budgets.
    DataModel(crate::workbook::data_model::LimitKind),
    /// External-link XML and cache budgets.
    ExternalLinks(crate::external_links::LimitKind),
    /// Worksheet OLE-object XML and payload budgets.
    OleObjects(crate::ole_objects::LimitKind),
    /// Workbook connection XML and collection budgets.
    Connections(crate::connections::LimitKind),
    /// Rich-value refresh interval XML and collection budgets.
    RichValueRefresh(crate::rich_values::refresh_intervals::LimitKind),
    /// Rich-value XML, collection, and relationship budgets.
    RichValues(crate::rich_values::LimitKind),
    /// Form-control XML, model, and retained-memory budgets.
    FormControl(crate::form_control::LimitKind),
    /// Source-backed tab-state XML audit budgets.
    TabState(crate::tab_state::source::LimitKind),
}

impl LimitKind {
    /// A short structural description of the dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        use crate::conditional_formatting::edit::LimitKind as ConditionalFormatting;
        use crate::drawing::edit::LimitKind as Drawing;
        match self {
            Self::Workbook(kind) => kind.as_str(),
            Self::ConditionalFormatting(ConditionalFormatting::Differentials) => {
                "conditional-formatting differential formats"
            },
            Self::ConditionalFormatting(ConditionalFormatting::Formats) => {
                "conditional-formatting collections of one worksheet"
            },
            Self::DataValidation(_) => "data validations of one worksheet",
            Self::AutoFilter(_) => "auto-filter columns",
            Self::Comments(crate::workbook::comments::edit::LimitKind::Comments) => {
                "notes of one worksheet"
            },
            Self::Comments(crate::workbook::comments::edit::LimitKind::TextBytes) => {
                "bytes of one note text or author"
            },
            Self::Chart(crate::chart::edit::LimitKind::Charts) => "charts of one worksheet",
            Self::Chart(crate::chart::edit::LimitKind::ChartBytes) => "bytes of one chart part",
            Self::Drawing(Drawing::Elements) => "objects of one worksheet drawing",
            Self::Drawing(Drawing::DrawingBytes) => "bytes of one drawing part",
            Self::Drawing(Drawing::ImageBytes) => "bytes of one image",
            Self::Drawing(_) => "worksheet drawing",
            Self::Opc(_) => "parts changed by one raw OPC commit",
            Self::CustomData(_) => "Custom Data budget",
            Self::Survey(_) => "Survey budget",
            Self::CalculationProperties(_) => "calculation-metadata budget",
            Self::PageBreaks(crate::page_breaks::LimitKind::HorizontalBreaks) => {
                "row breaks of one worksheet"
            },
            Self::PageBreaks(crate::page_breaks::LimitKind::VerticalBreaks) => {
                "column breaks of one worksheet"
            },
            Self::PivotTableData(kind) => kind.as_str(),
            Self::PivotServerFormats(kind) => kind.as_str(),
            Self::CachedUniqueNames(kind) => kind.as_str(),
            Self::XmlMaps(kind) => kind.as_str(),
            Self::CellValues(kind) => kind.as_str(),
            Self::RowVisibility(kind) => kind.as_str(),
            Self::Revisions(kind) => kind.as_str(),
            Self::DataTypeIcons(kind) => kind.as_str(),
            Self::DataModel(kind) => kind.as_str(),
            Self::ExternalLinks(kind) => kind.as_str(),
            Self::OleObjects(kind) => kind.as_str(),
            Self::Connections(kind) => kind.as_str(),
            Self::RichValueRefresh(kind) => kind.as_str(),
            Self::RichValues(kind) => kind.as_str(),
            Self::FormControl(kind) => kind.as_str(),
            Self::TabState(kind) => kind.as_str(),
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

macro_rules! limit_kind_from {
    ($($variant:ident($kind:ty)),* $(,)?) => {
        $(impl From<$kind> for LimitKind {
            fn from(kind: $kind) -> Self {
                Self::$variant(kind)
            }
        })*
    };
}

limit_kind_from!(
    Workbook(crate::workbook::edit::LimitKind),
    ConditionalFormatting(crate::conditional_formatting::edit::LimitKind),
    DataValidation(crate::data_validation::edit::LimitKind),
    AutoFilter(crate::auto_filter::edit::LimitKind),
    Comments(crate::workbook::comments::edit::LimitKind),
    Chart(crate::chart::edit::LimitKind),
    Drawing(crate::drawing::edit::LimitKind),
    Opc(crate::opc::LimitKind),
    CustomData(crate::custom_data::LimitKind),
    Survey(crate::survey::LimitKind),
    CalculationProperties(crate::calculation_properties::LimitKind),
    PageBreaks(crate::page_breaks::LimitKind),
    PivotTableData(crate::pivot::table_data::LimitKind),
);

impl From<std::convert::Infallible> for Error {
    fn from(never: std::convert::Infallible) -> Self {
        match never {}
    }
}

impl From<litchi_spreadsheet_drawing::Error> for Error {
    fn from(error: litchi_spreadsheet_drawing::Error) -> Self {
        match error {
            litchi_spreadsheet_drawing::Error::Drawing(error) => Self::Drawing(error),
            litchi_spreadsheet_drawing::Error::Mce(error) => Self::MarkupCompatibility(error),
            litchi_spreadsheet_drawing::Error::Limit {
                kind,
                observed,
                maximum,
            } => Self::limit(kind, observed, maximum),
            litchi_spreadsheet_drawing::Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => Self::InvalidLimit {
                kind: kind.into(),
                value: u64::try_from(value).unwrap_or(u64::MAX),
                maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            },
            litchi_spreadsheet_drawing::Error::Invalid(message)
            | litchi_spreadsheet_drawing::Error::Encoding(message) => Self::Invalid(message),
            other => Self::Invalid(other.to_string()),
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
    /// The requested merge intersects a worksheet-owned Excel table.
    TableOverlap {
        table: Rect,
    },
    Overlap {
        existing: Rect,
    },
    FollowerContent {
        address: Address,
    },
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
            Self::TableOverlap { table } => {
                write!(formatter, "it overlaps Excel table range {table}")
            },
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

/// A malformed-XML error from a parser or decoder failure.
pub(crate) fn malformed(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

/// Formatting into an owned `String` cannot fail; this maps the formatter's
/// error type for the writers that use `write!`.
pub(crate) fn fmt_error(_: std::fmt::Error) -> Error {
    invalid("formatting XML into a string failed")
}

pub(crate) fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use std::error::Error as _;

    use super::{Error, allocation};

    #[test]
    fn resource_refusal_retains_typed_budget_details() {
        let error = Error::from(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::Memory,
            observed: 65,
            limit: 64,
            scope: "pivot table retained cells".into(),
        });
        let Error::ResourceLimit(limit) = error else {
            panic!("budget refusal must remain distinguishable from invalid XML");
        };
        assert_eq!(limit.resource, litchi_core::Resource::Memory);
        assert_eq!(limit.observed, 65);
        assert_eq!(limit.limit, 64);
        assert_eq!(&*limit.scope, "pivot table retained cells");
    }

    #[test]
    fn form_control_refusal_retains_leaf_limit_details() {
        let error = Error::from(crate::form_control::FormControlError::Limit {
            kind: crate::form_control::LimitKind::Items,
            observed: 3,
            maximum: 2,
        });
        assert!(matches!(
            error,
            Error::FormControl(crate::form_control::FormControlError::Limit {
                kind: crate::form_control::LimitKind::Items,
                observed: 3,
                maximum: 2,
            })
        ));
    }

    #[test]
    fn form_control_execution_cancellation_remains_typed() {
        let leaf =
            crate::form_control::FormControlError::from(litchi_core::ExecutionError::Cancelled);
        assert!(matches!(
            Error::from(leaf),
            Error::FormControl(crate::form_control::FormControlError::Execution(
                litchi_core::ExecutionError::Cancelled
            ))
        ));
    }

    #[test]
    fn form_control_execution_limit_retains_resource_scope_and_full_width() {
        let scope: std::sync::Arc<str> = "caller/form-control-parser".into();
        let expected = litchi_core::ResourceLimit {
            resource: litchi_core::Resource::Work,
            observed: u64::MAX,
            limit: u64::MAX - 1,
            scope: scope.clone(),
        };
        let leaf = crate::form_control::FormControlError::from(
            litchi_core::ExecutionError::ResourceLimit(expected.clone()),
        );
        let Error::FormControl(crate::form_control::FormControlError::Execution(
            litchi_core::ExecutionError::ResourceLimit(actual),
        )) = Error::from(leaf)
        else {
            panic!("execution budget refusal must retain its typed source");
        };
        assert_eq!(actual, expected);
        assert!(std::sync::Arc::ptr_eq(&actual.scope, &scope));
    }

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
