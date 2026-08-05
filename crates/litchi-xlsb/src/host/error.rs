//! Error types for XLSB file parsing

use std::collections::TryReserveError;
use std::fmt;

/// Result type alias for XLSB operations
pub type Result<T> = std::result::Result<T, Error>;

/// Errors that can occur during XLSB file parsing
#[derive(Debug)]
pub enum Error {
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
    UnexpectedEndOfStream(String),
    /// Invalid formula
    InvalidFormula(String),
    /// Invalid cell reference
    InvalidCellReference(String),
    /// Worksheet not found
    WorksheetNotFound(String),
    /// File not found in ZIP
    FileNotFound(String),
    /// Unsupported feature
    UnsupportedFeature(String),
    /// Encoding error
    Encoding(String),
    /// Invalid package or DrawingML structure.
    InvalidFormat(String),
    /// Invalid OPC part URI.
    InvalidUri(String),
    /// Invalid relationship metadata.
    InvalidRelationship(String),
    /// Part content type did not match the XLSB graph contract.
    InvalidContentType {
        /// Expected content type.
        expected: String,
        /// Actual content type.
        got: String,
    },
    /// A bounded host-side operation could not reserve required memory.
    Allocation {
        /// Resource whose bounded plan could not be reserved.
        resource: &'static str,
        /// Original allocator failure.
        source: TryReserveError,
    },
    /// Shared DrawingML parsing error.
    Drawing(litchi_drawingml::Error),
    /// Shared host-neutral OOXML package-service error.
    Common(litchi_ooxml_common::Error),
    /// Bounded, inert VBA parsing or authoring error.
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
        typ: String,
        /// Value found
        val: String,
    },
    /// Workbook is password protected
    PasswordProtected,
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::Io(e) => write!(f, "I/O error: {}", e),
            Error::Xml(e) => write!(f, "XML error: {}", e),
            Error::Wire(e) => write!(f, "BIFF12 wire error: {e}"),
            Error::Calc(e) => write!(f, "XLSB calculation-property error: {e}"),
            Error::MergedCell(e) => write!(f, "XLSB merged-cell error: {e}"),
            Error::Hyperlink(e) => write!(f, "XLSB hyperlink error: {e}"),
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
            Error::WorksheetNotFound(name) => {
                write!(f, "Worksheet '{}' not found", name)
            },
            Error::FileNotFound(file) => {
                write!(f, "File '{}' not found in ZIP", file)
            },
            Error::UnsupportedFeature(feature) => {
                write!(f, "Unsupported feature: {}", feature)
            },
            Error::Encoding(msg) => {
                write!(f, "Encoding error: {}", msg)
            },
            Error::InvalidFormat(msg) => write!(f, "Invalid format: {msg}"),
            Error::InvalidUri(uri) => write!(f, "Invalid URI: {uri}"),
            Error::InvalidRelationship(msg) => write!(f, "Invalid relationship: {msg}"),
            Error::InvalidContentType { expected, got } => {
                write!(f, "Invalid content type: expected {expected}, got {got}")
            },
            Error::Allocation { resource, source } => {
                write!(f, "allocation failed for {resource}: {source}")
            },
            Error::Drawing(error) => write!(f, "DrawingML error: {error}"),
            Error::Common(error) => write!(f, "shared OOXML error: {error}"),
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
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Error::Io(e) => Some(e),
            Error::Xml(e) => Some(e),
            Error::Wire(e) => Some(e),
            Error::Calc(e) => Some(e),
            Error::MergedCell(e) => Some(e),
            Error::Hyperlink(e) => Some(e),
            Error::Allocation { source, .. } => Some(source),
            Error::Drawing(e) => Some(e),
            Error::Common(e) => Some(e),
            Error::Vba(e) => Some(e),
            _ => None,
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
    fn from(error: litchi_ooxml_common::XmlError) -> Self {
        Self::Encoding(error.to_string())
    }
}

impl From<litchi_ooxml_common::MceError> for Error {
    fn from(error: litchi_ooxml_common::MceError) -> Self {
        Self::Encoding(error.to_string())
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
    fn from(err: litchi_core::binary::BinaryError) -> Self {
        Error::Encoding(err.to_string())
    }
}

impl From<litchi_drawingml::Error> for Error {
    fn from(error: litchi_drawingml::Error) -> Self {
        Self::Drawing(error)
    }
}

impl From<litchi_ooxml_common::Error> for Error {
    fn from(error: litchi_ooxml_common::Error) -> Self {
        Self::Common(error)
    }
}

impl From<litchi_vba::Error> for Error {
    fn from(error: litchi_vba::Error) -> Self {
        Self::Vba(error)
    }
}

impl From<litchi_opc::error::OpcError> for Error {
    fn from(err: litchi_opc::error::OpcError) -> Self {
        Error::Encoding(format!("OPC error: {}", err))
    }
}

impl From<String> for Error {
    fn from(err: String) -> Self {
        Error::Encoding(err)
    }
}
