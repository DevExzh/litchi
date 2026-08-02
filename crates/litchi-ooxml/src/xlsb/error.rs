//! Error types for XLSB file parsing

use std::fmt;

/// Result type alias for XLSB operations
pub type XlsbResult<T> = Result<T, XlsbError>;

/// Errors that can occur during XLSB file parsing
#[derive(Debug)]
pub enum XlsbError {
    /// I/O error
    Io(std::io::Error),
    /// ZIP error
    Zip(String),
    /// XML parsing error
    Xml(quick_xml::Error),
    /// Validated BIFF12 wire-kernel error.
    Wire(litchi_xlsb::Error),
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

impl fmt::Display for XlsbError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            XlsbError::Io(e) => write!(f, "I/O error: {}", e),
            XlsbError::Zip(e) => write!(f, "ZIP error: {}", e),
            XlsbError::Xml(e) => write!(f, "XML error: {}", e),
            XlsbError::Wire(e) => write!(f, "BIFF12 wire error: {e}"),
            XlsbError::InvalidRecordType(rt) => write!(f, "Invalid record type: 0x{:04X}", rt),
            XlsbError::UnexpectedRecord { expected, found } => {
                write!(
                    f,
                    "Unexpected record type 0x{:04X}, expected 0x{:04X}",
                    found, expected
                )
            },
            XlsbError::InvalidLength { expected, found } => {
                write!(f, "Invalid length: expected {}, found {}", expected, found)
            },
            XlsbError::UnexpectedEndOfStream(context) => {
                write!(f, "Unexpected end of stream: {}", context)
            },
            XlsbError::InvalidFormula(msg) => {
                write!(f, "Invalid formula: {}", msg)
            },
            XlsbError::InvalidCellReference(ref_str) => {
                write!(f, "Invalid cell reference: {}", ref_str)
            },
            XlsbError::WorksheetNotFound(name) => {
                write!(f, "Worksheet '{}' not found", name)
            },
            XlsbError::FileNotFound(file) => {
                write!(f, "File '{}' not found in ZIP", file)
            },
            XlsbError::UnsupportedFeature(feature) => {
                write!(f, "Unsupported feature: {}", feature)
            },
            XlsbError::Encoding(msg) => {
                write!(f, "Encoding error: {}", msg)
            },
            XlsbError::Drawing(error) => write!(f, "DrawingML error: {error}"),
            XlsbError::Common(error) => write!(f, "shared OOXML error: {error}"),
            XlsbError::Vba(error) => write!(f, "VBA error: {error}"),
            XlsbError::WideStringLength { expected, actual } => {
                write!(
                    f,
                    "Wide string length mismatch: expected {}, actual {}",
                    expected, actual
                )
            },
            XlsbError::Unrecognized { typ, val } => {
                write!(f, "Unrecognized {}: {}", typ, val)
            },
            XlsbError::PasswordProtected => {
                write!(f, "Workbook is password protected")
            },
        }
    }
}

impl std::error::Error for XlsbError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            XlsbError::Io(e) => Some(e),
            XlsbError::Xml(e) => Some(e),
            XlsbError::Wire(e) => Some(e),
            XlsbError::Drawing(e) => Some(e),
            XlsbError::Common(e) => Some(e),
            XlsbError::Vba(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for XlsbError {
    fn from(err: std::io::Error) -> Self {
        XlsbError::Io(err)
    }
}

impl From<soapberry_zip::Error> for XlsbError {
    fn from(err: soapberry_zip::Error) -> Self {
        XlsbError::Zip(err.to_string())
    }
}

impl From<quick_xml::Error> for XlsbError {
    fn from(err: quick_xml::Error) -> Self {
        XlsbError::Xml(err)
    }
}

impl From<litchi_xlsb::Error> for XlsbError {
    fn from(error: litchi_xlsb::Error) -> Self {
        Self::Wire(error)
    }
}

impl From<litchi_core::binary::BinaryError> for XlsbError {
    fn from(err: litchi_core::binary::BinaryError) -> Self {
        XlsbError::Encoding(err.to_string())
    }
}

impl From<litchi_drawingml::Error> for XlsbError {
    fn from(error: litchi_drawingml::Error) -> Self {
        Self::Drawing(error)
    }
}

impl From<litchi_ooxml_common::Error> for XlsbError {
    fn from(error: litchi_ooxml_common::Error) -> Self {
        Self::Common(error)
    }
}

impl From<litchi_vba::Error> for XlsbError {
    fn from(error: litchi_vba::Error) -> Self {
        Self::Vba(error)
    }
}

#[cfg(feature = "encryption")]
impl From<litchi_cfb::OleError> for XlsbError {
    fn from(err: litchi_cfb::OleError) -> Self {
        match err {
            litchi_cfb::OleError::Io(e) => XlsbError::Io(e),
            litchi_cfb::OleError::InvalidFormat(msg) => {
                XlsbError::Encoding(format!("Invalid format: {}", msg))
            },
            litchi_cfb::OleError::InvalidData(msg) => {
                XlsbError::Encoding(format!("Invalid data: {}", msg))
            },
            litchi_cfb::OleError::NotOleFile => XlsbError::Encoding("Not an OLE file".to_string()),
            litchi_cfb::OleError::CorruptedFile(msg) => {
                XlsbError::Encoding(format!("Corrupted file: {}", msg))
            },
            litchi_cfb::OleError::StreamNotFound => {
                XlsbError::FileNotFound("Stream not found".to_string())
            },
        }
    }
}

impl From<litchi_opc::error::OpcError> for XlsbError {
    fn from(err: litchi_opc::error::OpcError) -> Self {
        XlsbError::Encoding(format!("OPC error: {}", err))
    }
}

impl From<crate::error::OoxmlError> for XlsbError {
    fn from(err: crate::error::OoxmlError) -> Self {
        match err {
            crate::error::OoxmlError::Opc(e) => XlsbError::Encoding(format!("OPC error: {}", e)),
            crate::error::OoxmlError::Xml(msg) => {
                XlsbError::Encoding(format!("XML error: {}", msg))
            },
            crate::error::OoxmlError::PartNotFound(path) => XlsbError::FileNotFound(path),
            crate::error::OoxmlError::InvalidContentType { expected, got } => XlsbError::Encoding(
                format!("Invalid content type: expected {}, got {}", expected, got),
            ),
            crate::error::OoxmlError::InvalidUri(s) => {
                XlsbError::Encoding(format!("Invalid URI: {}", s))
            },
            crate::error::OoxmlError::InvalidRelationship(msg) => {
                XlsbError::Encoding(format!("Invalid relationship: {}", msg))
            },
            crate::error::OoxmlError::InvalidFormat(msg) => XlsbError::Encoding(msg),
            crate::error::OoxmlError::Drawing(err) => XlsbError::Encoding(err.to_string()),
            crate::error::OoxmlError::Docx(err) => XlsbError::Encoding(err.to_string()),
            crate::error::OoxmlError::Pptx(err) => XlsbError::Encoding(err.to_string()),
            crate::error::OoxmlError::Xlsb(err) => XlsbError::Wire(err),
            crate::error::OoxmlError::Io(e) => XlsbError::Io(e),
            crate::error::OoxmlError::Common(err) => XlsbError::Common(err),
            crate::error::OoxmlError::Vba(err) => XlsbError::Vba(err),
            crate::error::OoxmlError::UnsafeEdit {
                format,
                operation,
                reason,
            } => XlsbError::UnsupportedFeature(format!(
                "unsafe {format} edit rejected during {operation}: {reason}"
            )),
            crate::error::OoxmlError::Other(msg) => XlsbError::Encoding(msg),
        }
    }
}

impl From<String> for XlsbError {
    fn from(err: String) -> Self {
        XlsbError::Encoding(err)
    }
}
