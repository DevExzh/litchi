//! Error types for XLSB file parsing

#![deny(
    clippy::cast_lossless,
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::checked_conversions,
    clippy::expect_used,
    clippy::float_cmp,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::unnecessary_unwrap,
    clippy::wildcard_enum_match_arm,
    reason = "typed package errors contain no unchecked BIFF12 conversions or panic extraction"
)]

use std::collections::TryReserveError;
use std::fmt;

/// Result type alias for XLSB operations
pub type Result<T> = std::result::Result<T, Error>;

/// Errors that can occur during XLSB file parsing
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// A textual sheet-name dependency could not be rewritten safely.
    SheetNameDependency(litchi_ooxml_common::xml::sheet_name::Error),
    /// I/O error
    Io(std::io::Error),
    /// XML parsing error
    Xml(quick_xml::Error),
    /// Validated BIFF12 wire-kernel error.
    Wire(crate::Error),
    /// Typed workbook calculation-property error.
    Calc(crate::calc::Error),
    /// Typed merged-cell record error.
    MergedCell(crate::merged_cells::Error),
    /// Typed hyperlink record error.
    Hyperlink(crate::hyperlinks::Error),
    /// Typed sparkline record error.
    Sparkline(crate::sparkline::Error),
    /// Invalid checked cell coordinate supplied by a caller.
    Coordinate(litchi_sheet::CoordinateError),
    /// Invalid checked cell rectangle supplied by a caller.
    Range(litchi_sheet::RangeError),
    /// Invalid record type
    InvalidRecordType(u16),
    /// Unexpected record
    UnexpectedRecord {
        /// Expected record type
        expected: u16,
        /// Found record type
        found: u16,
    },
    /// Invalid data length
    InvalidLength {
        /// Expected length
        expected: usize,
        /// Found length
        found: usize,
    },
    /// End of stream reached unexpectedly
    UnexpectedEndOfStream(&'static str),
    /// Invalid formula
    InvalidFormula(&'static str),
    /// Invalid cell reference
    InvalidCellReference(&'static str),
    /// No worksheet matches the selector: the zero-based position when the
    /// selector was positional, `None` for a name (which is not echoed).
    WorksheetNotFound(Option<usize>),
    /// File not found in ZIP
    FileNotFound(&'static str),
    /// Unsupported feature
    UnsupportedFeature(&'static str),
    /// Formula text requires workbook metadata that is missing or ambiguous.
    UnresolvedDependency(&'static str),
    /// A drawing transfer was refused before any package mutation.
    DrawingTransfer(crate::cell_values::DrawingTransferRefusal),
    /// Encoding error
    Encoding(&'static str),
    /// Invalid package or DrawingML structure.
    InvalidFormat(&'static str),
    /// Invalid OPC part URI.
    InvalidUri(&'static str),
    /// Invalid relationship metadata.
    InvalidRelationship(&'static str),
    /// Part content type did not match the XLSB graph contract.
    InvalidContentType {
        /// Expected content type; the actual one is not echoed.
        expected: &'static str,
    },
    /// A bounded host-side operation could not reserve required memory.
    Allocation {
        /// Resource whose bounded plan could not be reserved.
        resource: &'static str,
        /// Original allocator failure.
        source: TryReserveError,
    },
    /// Aggregate capacity arithmetic overflowed before allocation.
    CapacityOverflow {
        /// Resource whose aggregate capacity could not be represented.
        resource: &'static str,
    },
    /// A governed external-link resource limit was exceeded.
    ExternalLinkLimitExceeded {
        /// External-link resource whose configured ceiling was crossed.
        resource: crate::external_link::ExternalLinkResource,
        /// Observed or requested amount.
        actual: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// Shared DrawingML parsing error.
    Drawing(litchi_drawingml::Error),
    /// Shared host-neutral OOXML package-service error.
    Common(litchi_ooxml_common::Error),
    /// Deferred OPC package, source-freshness, or resource-limit error.
    Opc(litchi_opc::OpcError),
    /// Cooperative cancellation or caller execution-budget refusal.
    Execution(litchi_core::ExecutionError),
    /// A source-checked patch was applied to a different owner closure or
    /// byte image.
    PatchConflict {
        /// Human-readable closure name.
        part: &'static str,
    },
    /// A source-backed package changed while the operation was in progress.
    SourceChanged {
        /// Version captured before the operation.
        expected: litchi_core::SourceVersion,
        /// Version observed when the operation completed.
        observed: litchi_core::SourceVersion,
    },
    /// A generated staging or publication invariant was not satisfied.
    Internal(&'static str),
    /// The package contains a signature/edit policy that forbids this change.
    Signed,
    /// A recognized Custom Data storage still has inbound BIFF12 references.
    CustomDataReferenced {
        /// Catalog position of the storage.
        storage: usize,
        /// Number of recognized references.
        connections: usize,
    },
    /// Bounded, inert VBA parsing or authoring error.
    #[cfg(feature = "vba-inspection")]
    Vba(litchi_vba::Error),
    /// Wide string length error
    WideStringLength {
        /// Expected length
        expected: usize,
        /// Actual length
        actual: usize,
    },
    /// Unrecognized data
    Unrecognized {
        /// Data type
        typ: &'static str,
        /// Value found
        val: &'static str,
    },
    /// Workbook is password protected
    PasswordProtected,
    /// A well-formed request this route refuses (record 0351 §2.4),
    /// classified by its [`litchi_core::edit::RefusalKind`].
    Refused(litchi_core::edit::Refusal),
    /// A configured budget was exceeded.
    Limit {
        /// The bounded dimension.
        kind: LimitKind,
        /// The observed (or first refused) amount.
        observed: usize,
        /// The configured maximum.
        maximum: usize,
    },
    /// An open route was given a policy that its owner kind cannot honour.
    UnsupportedOpenPolicy {
        /// The refused policy.
        policy: crate::OpenPolicy,
    },
    /// A configured limit is not valid.
    InvalidLimit {
        /// The configured dimension.
        kind: LimitKind,
        /// The configured value.
        value: usize,
        /// The largest value the dimension admits.
        maximum: usize,
    },
}

/// A bounded dimension of an XLSB package operation.
///
/// Each variant wraps the typed `LimitKind` of the family that owns the
/// bound.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Cell-value and primary workbook transactions
    /// ([`crate::cell_values::LimitKind`]).
    CellValues(crate::cell_values::LimitKind),
    /// Raw OPC edits ([`crate::opc::LimitKind`]).
    Opc(crate::opc::LimitKind),
    /// Worksheet binary indexes ([`crate::binary_index::LimitKind`]).
    BinaryIndex(crate::binary_index::LimitKind),
    /// The Calculation Chain family ([`crate::calculation_chain::LimitKind`]).
    CalculationChain(crate::calculation_chain::LimitKind),
    /// The External Data Connections family ([`crate::package::connections::LimitKind`]).
    Connections(crate::package::connections::LimitKind),
    /// The Custom Data family ([`crate::custom_data::LimitKind`]).
    CustomData(crate::custom_data::LimitKind),
    /// The Data Model family ([`crate::data_model::LimitKind`]).
    DataModel(crate::data_model::LimitKind),
    /// The Theme family ([`crate::theme::LimitKind`]).
    Theme(crate::theme::LimitKind),
    /// The Volatile Dependencies family ([`crate::volatile_dependencies::LimitKind`]).
    VolatileDependencies(crate::volatile_dependencies::LimitKind),
    /// Source-backed workbook indexing ([`crate::workbook::LimitKind`]).
    Workbook(crate::workbook::LimitKind),
    /// Worksheet-image authoring ([`crate::package::drawing_image::LimitKind`]).
    Image(crate::package::drawing_image::LimitKind),
    /// Shared SpreadsheetDrawing XML scanners and rewriters.
    Drawing(litchi_spreadsheet_drawing::LimitKind),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::CellValues(kind) => fmt::Display::fmt(kind, f),
            Self::Opc(crate::opc::LimitKind::Parts) => f.write_str("changed package parts"),
            Self::Opc(crate::opc::LimitKind::PatchBytes) => f.write_str("retained patch bytes"),
            Self::BinaryIndex(kind) => fmt::Display::fmt(kind, f),
            Self::CalculationChain(kind) => fmt::Display::fmt(kind, f),
            Self::Connections(kind) => fmt::Display::fmt(kind, f),
            Self::CustomData(kind) => fmt::Display::fmt(kind, f),
            Self::DataModel(kind) => fmt::Display::fmt(kind, f),
            Self::Theme(kind) => fmt::Display::fmt(kind, f),
            Self::VolatileDependencies(kind) => fmt::Display::fmt(kind, f),
            Self::Workbook(kind) => fmt::Display::fmt(kind, f),
            Self::Image(kind) => fmt::Display::fmt(kind, f),
            Self::Drawing(kind) => fmt::Display::fmt(kind, f),
        }
    }
}

impl From<crate::cell_values::LimitKind> for LimitKind {
    fn from(kind: crate::cell_values::LimitKind) -> Self {
        Self::CellValues(kind)
    }
}

impl From<crate::binary_index::LimitKind> for LimitKind {
    fn from(kind: crate::binary_index::LimitKind) -> Self {
        Self::BinaryIndex(kind)
    }
}

impl From<crate::calculation_chain::LimitKind> for LimitKind {
    fn from(kind: crate::calculation_chain::LimitKind) -> Self {
        Self::CalculationChain(kind)
    }
}

impl From<crate::package::connections::LimitKind> for LimitKind {
    fn from(kind: crate::package::connections::LimitKind) -> Self {
        Self::Connections(kind)
    }
}

impl From<crate::custom_data::LimitKind> for LimitKind {
    fn from(kind: crate::custom_data::LimitKind) -> Self {
        Self::CustomData(kind)
    }
}

impl From<crate::data_model::LimitKind> for LimitKind {
    fn from(kind: crate::data_model::LimitKind) -> Self {
        Self::DataModel(kind)
    }
}

impl From<crate::theme::LimitKind> for LimitKind {
    fn from(kind: crate::theme::LimitKind) -> Self {
        Self::Theme(kind)
    }
}

impl From<crate::volatile_dependencies::LimitKind> for LimitKind {
    fn from(kind: crate::volatile_dependencies::LimitKind) -> Self {
        Self::VolatileDependencies(kind)
    }
}

impl From<crate::workbook::LimitKind> for LimitKind {
    fn from(kind: crate::workbook::LimitKind) -> Self {
        Self::Workbook(kind)
    }
}

impl From<crate::opc::LimitKind> for LimitKind {
    fn from(kind: crate::opc::LimitKind) -> Self {
        Self::Opc(kind)
    }
}

impl From<litchi_spreadsheet_drawing::LimitKind> for LimitKind {
    fn from(kind: litchi_spreadsheet_drawing::LimitKind) -> Self {
        Self::Drawing(kind)
    }
}

impl Error {
    /// A typed budget refusal.
    pub(crate) fn limit(kind: impl Into<LimitKind>, observed: usize, maximum: usize) -> Self {
        Self::Limit {
            kind: kind.into(),
            observed,
            maximum,
        }
    }

    /// A content-free refusal of a well-formed request.
    pub(crate) const fn refused(
        kind: litchi_core::edit::RefusalKind,
        operation: &'static str,
        reason: &'static str,
    ) -> Self {
        Self::Refused(litchi_core::edit::Refusal::new(kind, operation, reason))
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::SheetNameDependency(error) => fmt::Display::fmt(error, f),
            Error::Io(e) => write!(f, "I/O error: {}", e),
            Error::Xml(e) => write!(f, "XML error: {}", e),
            Error::Wire(e) => write!(f, "BIFF12 wire error: {e}"),
            Error::Calc(e) => write!(f, "XLSB calculation-property error: {e}"),
            Error::MergedCell(e) => write!(f, "XLSB merged-cell error: {e}"),
            Error::Hyperlink(e) => write!(f, "XLSB hyperlink error: {e}"),
            Error::Sparkline(e) => write!(f, "XLSB sparkline error: {e}"),
            Error::Coordinate(e) => write!(f, "XLSB coordinate error: {e}"),
            Error::Range(e) => write!(f, "XLSB range error: {e}"),
            Error::InvalidRecordType(rt) => write!(f, "Invalid record type: 0x{:04X}", rt),
            Error::UnexpectedRecord { expected, found } => {
                write!(
                    f,
                    "Unexpected record type 0x{:04X}, expected 0x{:04X}",
                    found, expected
                )
            },
            Error::InvalidLength { expected, found } => {
                write!(f, "Invalid length: expected {}, found {}", expected, found)
            },
            Error::UnexpectedEndOfStream(context) => {
                write!(f, "Unexpected end of stream: {}", context)
            },
            Error::InvalidFormula(msg) => {
                write!(f, "Invalid formula: {}", msg)
            },
            Error::InvalidCellReference(ref_str) => {
                write!(f, "Invalid cell reference: {}", ref_str)
            },
            Error::WorksheetNotFound(Some(position)) => {
                write!(f, "no worksheet at position {position}")
            },
            Error::WorksheetNotFound(None) => write!(f, "no worksheet has the requested name"),
            Error::FileNotFound(part) => write!(f, "the {part} part is missing from the package"),
            Error::UnsupportedFeature(feature) => {
                write!(f, "Unsupported feature: {}", feature)
            },
            Error::UnresolvedDependency(dependency) => {
                write!(f, "Unresolved formula dependency: {}", dependency)
            },
            Error::DrawingTransfer(refusal) => write!(f, "Drawing transfer refused: {refusal}"),
            Error::Encoding(msg) => {
                write!(f, "Encoding error: {}", msg)
            },
            Error::InvalidFormat(msg) => write!(f, "Invalid format: {msg}"),
            Error::InvalidUri(uri) => write!(f, "Invalid URI: {uri}"),
            Error::InvalidRelationship(msg) => write!(f, "Invalid relationship: {msg}"),
            Error::InvalidContentType { expected } => {
                write!(f, "a part does not have the content type {expected}")
            },
            Error::Allocation { resource, source } => {
                write!(f, "allocation failed for {resource}: {source}")
            },
            Error::CapacityOverflow { resource } => {
                write!(f, "capacity overflow while planning {resource}")
            },
            Error::ExternalLinkLimitExceeded {
                resource,
                actual,
                maximum,
            } => write!(
                f,
                "external-link {resource} limit exceeded: actual {actual}, maximum {maximum}"
            ),
            Error::Drawing(error) => write!(f, "DrawingML error: {error}"),
            Error::Common(error) => write!(f, "shared OOXML error: {error}"),
            Error::Opc(error) => write!(f, "OPC source error: {error}"),
            Error::Execution(error) => write!(f, "execution policy error: {error}"),
            Error::PatchConflict { part } => write!(f, "source patch conflict: {part}"),
            Error::SourceChanged { .. } => {
                write!(f, "XLSB source changed while the operation was in progress")
            },
            Error::Internal(reason) => write!(f, "XLSB internal invariant failed: {reason}"),
            Error::Signed => write!(f, "package signature/edit policy forbids this change"),
            Error::CustomDataReferenced {
                storage,
                connections,
            } => write!(
                f,
                "Custom Data storage {storage} is referenced by {connections} connection(s)"
            ),
            #[cfg(feature = "vba-inspection")]
            Error::Vba(error) => write!(f, "VBA error: {error}"),
            Error::WideStringLength { expected, actual } => {
                write!(
                    f,
                    "Wide string length mismatch: expected {}, actual {}",
                    expected, actual
                )
            },
            Error::Unrecognized { typ, val } => {
                write!(f, "Unrecognized {}: {}", typ, val)
            },
            Error::PasswordProtected => {
                write!(f, "Workbook is password protected")
            },
            Error::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                f,
                "XLSB {kind} limit exceeded: {observed} exceeds the maximum {maximum}"
            ),
            Error::UnsupportedOpenPolicy { policy } => {
                write!(f, "this XLSB open route does not support {policy}")
            },
            Error::Refused(refusal) => fmt::Display::fmt(refusal, f),
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                f,
                "invalid XLSB {kind} limit {value}: the maximum is {maximum}"
            ),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Error::SheetNameDependency(error) => Some(error),
            Error::Io(e) => Some(e),
            Error::Xml(e) => Some(e),
            Error::Wire(e) => Some(e),
            Error::Calc(e) => Some(e),
            Error::MergedCell(e) => Some(e),
            Error::Hyperlink(e) => Some(e),
            Error::Sparkline(e) => Some(e),
            Error::Coordinate(e) => Some(e),
            Error::Range(e) => Some(e),
            Error::DrawingTransfer(refusal) => Some(refusal),
            Error::Allocation { source, .. } => Some(source),
            Error::Drawing(e) => Some(e),
            Error::Common(e) => Some(e),
            Error::Opc(e) => Some(e),
            Error::Execution(e) => Some(e),
            #[cfg(feature = "vba-inspection")]
            Error::Vba(e) => Some(e),
            Error::InvalidRecordType(_)
            | Error::UnexpectedRecord { .. }
            | Error::InvalidLength { .. }
            | Error::UnexpectedEndOfStream(_)
            | Error::InvalidFormula(_)
            | Error::InvalidCellReference(_)
            | Error::WorksheetNotFound(_)
            | Error::FileNotFound(_)
            | Error::UnsupportedFeature(_)
            | Error::UnresolvedDependency(_)
            | Error::Encoding(_)
            | Error::InvalidFormat(_)
            | Error::InvalidUri(_)
            | Error::InvalidRelationship(_)
            | Error::InvalidContentType { .. }
            | Error::CapacityOverflow { .. }
            | Error::ExternalLinkLimitExceeded { .. }
            | Error::WideStringLength { .. }
            | Error::Unrecognized { .. }
            | Error::PasswordProtected
            | Error::Limit { .. }
            | Error::InvalidLimit { .. }
            | Error::Refused(_)
            | Error::UnsupportedOpenPolicy { .. } => None,
            Error::PatchConflict { .. }
            | Error::SourceChanged { .. }
            | Error::Internal(_)
            | Error::Signed
            | Error::CustomDataReferenced { .. } => None,
        }
    }
}

impl From<std::io::Error> for Error {
    fn from(err: std::io::Error) -> Self {
        Error::Io(err)
    }
}

impl From<quick_xml::Error> for Error {
    fn from(err: quick_xml::Error) -> Self {
        Error::Xml(err)
    }
}

impl From<litchi_ooxml_common::XmlError> for Error {
    fn from(_error: litchi_ooxml_common::XmlError) -> Self {
        Self::Encoding("the text cannot be decoded")
    }
}

impl From<litchi_ooxml_common::mce::Error> for Error {
    fn from(_error: litchi_ooxml_common::mce::Error) -> Self {
        Self::Encoding("the text cannot be decoded")
    }
}

impl From<crate::Error> for Error {
    fn from(error: crate::Error) -> Self {
        Self::Wire(error)
    }
}

impl From<crate::calc::Error> for Error {
    fn from(error: crate::calc::Error) -> Self {
        Self::Calc(error)
    }
}

impl From<crate::merged_cells::Error> for Error {
    fn from(error: crate::merged_cells::Error) -> Self {
        Self::MergedCell(error)
    }
}

impl From<crate::hyperlinks::Error> for Error {
    fn from(error: crate::hyperlinks::Error) -> Self {
        Self::Hyperlink(error)
    }
}

impl From<crate::comments::Error> for Error {
    fn from(error: crate::comments::Error) -> Self {
        match error {
            crate::comments::Error::Wire(error) => Self::Wire(error),
            crate::comments::Error::InvalidRecordType(record_type) => {
                Self::InvalidRecordType(record_type)
            },
            crate::comments::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            crate::comments::Error::Unrecognized { typ, val } => Self::Unrecognized { typ, val },
            crate::comments::Error::Encoding(message) => Self::Encoding(message),
            crate::comments::Error::UnsupportedFeature(feature) => {
                Self::UnsupportedFeature(feature)
            },
            crate::comments::Error::Allocation { resource, source } => {
                Self::Allocation { resource, source }
            },
        }
    }
}

impl From<crate::external_link::Error> for Error {
    fn from(error: crate::external_link::Error) -> Self {
        match error {
            crate::external_link::Error::Wire(error) => Self::Wire(error),
            crate::external_link::Error::InvalidFormula(message) => Self::InvalidFormula(message),
            crate::external_link::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            crate::external_link::Error::Allocation { resource, source } => {
                Self::Allocation { resource, source }
            },
            crate::external_link::Error::LimitExceeded {
                resource,
                actual,
                maximum,
            } => Self::ExternalLinkLimitExceeded {
                resource,
                actual,
                maximum,
            },
            crate::external_link::Error::PatchConflict => Self::PatchConflict {
                part: "external links",
            },
        }
    }
}

impl From<crate::conditional_formatting::Error> for Error {
    fn from(error: crate::conditional_formatting::Error) -> Self {
        match error {
            crate::conditional_formatting::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            crate::conditional_formatting::Error::InvalidFormula(message) => {
                Self::InvalidFormula(message)
            },
            crate::conditional_formatting::Error::InvalidCellReference(reference) => {
                Self::InvalidCellReference(reference)
            },
            crate::conditional_formatting::Error::Encoding(message) => Self::Encoding(message),
            crate::conditional_formatting::Error::UnsupportedFeature(feature) => {
                Self::UnsupportedFeature(feature)
            },
            crate::conditional_formatting::Error::Unrecognized { typ, val } => {
                Self::Unrecognized { typ, val }
            },
            crate::conditional_formatting::Error::Wire(error) => Self::from(error),
            crate::conditional_formatting::Error::Formula(error) => Self::from(error),
            crate::conditional_formatting::Error::Io(error) => Self::Io(error),
        }
    }
}

impl From<litchi_core::binary::BinaryError> for Error {
    fn from(_err: litchi_core::binary::BinaryError) -> Self {
        Error::Encoding("the text cannot be decoded")
    }
}

impl From<litchi_drawingml::Error> for Error {
    fn from(error: litchi_drawingml::Error) -> Self {
        Self::Drawing(error)
    }
}

impl From<litchi_spreadsheet_drawing::Error> for Error {
    fn from(error: litchi_spreadsheet_drawing::Error) -> Self {
        match error {
            litchi_spreadsheet_drawing::Error::Drawing(error) => Self::Drawing(error),
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
                value,
                maximum,
            },
            litchi_spreadsheet_drawing::Error::Invalid(_) => {
                Self::InvalidFormat("the drawing is not valid")
            },
            litchi_spreadsheet_drawing::Error::Encoding(_) => {
                Self::Encoding("the drawing text cannot be decoded")
            },
            _other @ (litchi_spreadsheet_drawing::Error::Mce(_)
            | litchi_spreadsheet_drawing::Error::Xml(_)
            | _) => Self::InvalidFormat("the package structure is not valid"),
        }
    }
}

impl From<litchi_ooxml_common::Error> for Error {
    fn from(error: litchi_ooxml_common::Error) -> Self {
        match error {
            litchi_ooxml_common::Error::Refused(refusal) => Self::Refused(refusal),
            litchi_ooxml_common::Error::PatchConflict => Self::PatchConflict {
                part: "shared OOXML graph",
            },
            litchi_ooxml_common::Error::Signed => Self::Signed,
            litchi_ooxml_common::Error::SourceChanged { expected, observed } => {
                Self::SourceChanged { expected, observed }
            },
            litchi_ooxml_common::Error::Internal(reason) => Self::Internal(reason),
            other @ litchi_ooxml_common::Error::Opc(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::Patch(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::Xml(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::Missing(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::ContentType { .. } => Self::Common(other),
            other @ litchi_ooxml_common::Error::Relationship(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::Invalid(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::SpreadsheetXmlMaps(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::Limit { .. } => Self::Common(other),
            other @ litchi_ooxml_common::Error::Allocation { .. } => Self::Common(other),
            other @ litchi_ooxml_common::Error::Uri(_) => Self::Common(other),
            #[cfg(feature = "vba-inspection")]
            other @ litchi_ooxml_common::Error::Vba(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::Mce(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::Decode(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::Io(_) => Self::Common(other),
            other @ litchi_ooxml_common::Error::Fmt(_) => Self::Common(other),
            other => Self::Common(other),
        }
    }
}

#[cfg(feature = "vba-inspection")]
impl From<litchi_vba::Error> for Error {
    fn from(error: litchi_vba::Error) -> Self {
        Self::Vba(error)
    }
}

impl From<litchi_opc::error::OpcError> for Error {
    fn from(err: litchi_opc::error::OpcError) -> Self {
        Error::Opc(err)
    }
}

impl From<litchi_core::ExecutionError> for Error {
    fn from(error: litchi_core::ExecutionError) -> Self {
        Self::Execution(error)
    }
}

impl From<String> for Error {
    fn from(_err: String) -> Self {
        Error::Encoding("the text cannot be decoded")
    }
}

impl From<litchi_sheet::CoordinateError> for Error {
    fn from(error: litchi_sheet::CoordinateError) -> Self {
        Self::Coordinate(error)
    }
}

impl From<litchi_sheet::RangeError> for Error {
    fn from(error: litchi_sheet::RangeError) -> Self {
        Self::Range(error)
    }
}

impl From<litchi_ooxml_common::xml::sheet_name::Error> for Error {
    fn from(error: litchi_ooxml_common::xml::sheet_name::Error) -> Self {
        Self::SheetNameDependency(error)
    }
}

/// Classify an XLSB package error for the format-neutral sheet error the
/// formula engine reports. The message is static: no sheet name, reference
/// or record content crosses.
impl From<Error> for litchi_core::sheet::Error {
    #[cold]
    #[inline(never)]
    fn from(error: Error) -> Self {
        use litchi_core::edit::RefusalKind;
        let refused = |kind, reason| Self::refused(kind, "read", reason);
        match error {
            Error::Io(_) => Self::Io,
            Error::WorksheetNotFound(_) | Error::FileNotFound(_) => refused(
                RefusalKind::NotFound,
                "the XLSB workbook has no such worksheet",
            ),
            Error::InvalidCellReference(_) | Error::Coordinate(_) | Error::Range(_) => {
                refused(RefusalKind::InvalidValue, "the cell reference is not valid")
            },
            Error::UnsupportedFeature(_) => refused(
                RefusalKind::UnsupportedContent,
                "the XLSB workbook uses an unsupported feature",
            ),
            Error::Refused(refusal) => Self::Refused(refusal),
            Error::Limit { .. }
            | Error::InvalidLimit { .. }
            | Error::ExternalLinkLimitExceeded { .. }
            | Error::Allocation { .. }
            | Error::CapacityOverflow { .. } => refused(
                RefusalKind::UnsupportedContent,
                "the XLSB workbook exceeds a configured limit",
            ),
            Error::PasswordProtected | Error::Signed => {
                refused(RefusalKind::Protected, "the XLSB workbook is protected")
            },
            Error::UnsupportedOpenPolicy { .. } => refused(
                RefusalKind::UnsupportedOperation,
                "the XLSB workbook was opened without this capability",
            ),
            Error::Opc(_) | Error::Execution(_) => Self::Io,
            Error::SourceChanged { .. } => Self::Io,
            Error::Internal(reason) => Self::Internal(reason),
            Error::PatchConflict { .. } => Self::PatchConflict,
            Error::SheetNameDependency(_)
            | Error::Xml(_)
            | Error::Wire(_)
            | Error::Calc(_)
            | Error::MergedCell(_)
            | Error::Hyperlink(_)
            | Error::Sparkline(_)
            | Error::InvalidRecordType(_)
            | Error::UnexpectedRecord { .. }
            | Error::InvalidLength { .. }
            | Error::UnexpectedEndOfStream(_)
            | Error::InvalidFormula(_)
            | Error::UnresolvedDependency(_)
            | Error::DrawingTransfer(_)
            | Error::Encoding(_)
            | Error::InvalidFormat(_)
            | Error::InvalidUri(_)
            | Error::InvalidRelationship(_)
            | Error::InvalidContentType { .. }
            | Error::Drawing(_)
            | Error::Common(_)
            | Error::CustomDataReferenced { .. }
            | Error::WideStringLength { .. }
            | Error::Unrecognized { .. } => Self::Malformed("the XLSB data is malformed"),
            #[cfg(feature = "vba-inspection")]
            Error::Vba(_) => Self::Malformed("the XLSB data is malformed"),
        }
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn sheet_conversion_preserves_patch_conflict() {
        let error = super::Error::PatchConflict { part: "workbook" };
        assert!(matches!(
            litchi_core::sheet::Error::from(error),
            litchi_core::sheet::Error::PatchConflict
        ));
    }
}
