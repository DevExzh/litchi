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
    Zip(zip::result::ZipError),
    /// XML parsing error
    Xml(quick_xml::Error),
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
            XlsbError::InvalidRecordType(rt) => write!(f, "Invalid record type: 0x{:04X}", rt),
            XlsbError::UnexpectedRecord { expected, found } => {
                write!(f, "Unexpected record type 0x{:04X}, expected 0x{:04X}", found, expected)
            }
            XlsbError::InvalidLength { expected, found } => {
                write!(f, "Invalid length: expected {}, found {}", expected, found)
            }
            XlsbError::UnexpectedEndOfStream(context) => {
                write!(f, "Unexpected end of stream: {}", context)
            }
            XlsbError::InvalidFormula(msg) => {
                write!(f, "Invalid formula: {}", msg)
            }
            XlsbError::InvalidCellReference(ref_str) => {
                write!(f, "Invalid cell reference: {}", ref_str)
            }
            XlsbError::WorksheetNotFound(name) => {
                write!(f, "Worksheet '{}' not found", name)
            }
            XlsbError::FileNotFound(file) => {
                write!(f, "File '{}' not found in ZIP", file)
            }
            XlsbError::UnsupportedFeature(feature) => {
                write!(f, "Unsupported feature: {}", feature)
            }
            XlsbError::Encoding(msg) => {
                write!(f, "Encoding error: {}", msg)
            }
            XlsbError::WideStringLength { expected, actual } => {
                write!(f, "Wide string length mismatch: expected {}, actual {}", expected, actual)
            }
            XlsbError::Unrecognized { typ, val } => {
                write!(f, "Unrecognized {}: {}", typ, val)
            }
            XlsbError::PasswordProtected => {
                write!(f, "Workbook is password protected")
            }
        }
    }
}

impl std::error::Error for XlsbError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            XlsbError::Io(e) => Some(e),
            XlsbError::Zip(e) => Some(e),
            XlsbError::Xml(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for XlsbError {
    fn from(err: std::io::Error) -> Self {
        XlsbError::Io(err)
    }
}

impl From<zip::result::ZipError> for XlsbError {
    fn from(err: zip::result::ZipError) -> Self {
        XlsbError::Zip(err)
    }
}

impl From<quick_xml::Error> for XlsbError {
    fn from(err: quick_xml::Error) -> Self {
        XlsbError::Xml(err)
    }
}

impl From<crate::common::binary::BinaryError> for XlsbError {
    fn from(err: crate::common::binary::BinaryError) -> Self {
        XlsbError::Encoding(err.to_string())
    }
}

#[cfg(feature = "ole")]
impl From<crate::ole::OleError> for XlsbError {
    fn from(err: crate::ole::OleError) -> Self {
        match err {
            crate::ole::OleError::Io(e) => XlsbError::Io(e),
            crate::ole::OleError::InvalidFormat(msg) => XlsbError::Encoding(format!("Invalid format: {}", msg)),
            crate::ole::OleError::InvalidData(msg) => XlsbError::Encoding(format!("Invalid data: {}", msg)),
            crate::ole::OleError::NotOleFile => XlsbError::Encoding("Not an OLE file".to_string()),
            crate::ole::OleError::CorruptedFile(msg) => XlsbError::Encoding(format!("Corrupted file: {}", msg)),
            crate::ole::OleError::StreamNotFound => XlsbError::FileNotFound("Stream not found".to_string()),
        }
    }
}

impl From<crate::ooxml::opc::error::OpcError> for XlsbError {
    fn from(err: crate::ooxml::opc::error::OpcError) -> Self {
        XlsbError::Encoding(format!("OPC error: {}", err))
    }
}

impl From<crate::ooxml::error::OoxmlError> for XlsbError {
    fn from(err: crate::ooxml::error::OoxmlError) -> Self {
        match err {
            crate::ooxml::error::OoxmlError::Opc(e) => XlsbError::Encoding(format!("OPC error: {}", e)),
            crate::ooxml::error::OoxmlError::Xml(msg) => XlsbError::Encoding(format!("XML error: {}", msg)),
            crate::ooxml::error::OoxmlError::PartNotFound(path) => XlsbError::FileNotFound(path),
            crate::ooxml::error::OoxmlError::InvalidContentType { expected, got } => XlsbError::Encoding(format!("Invalid content type: expected {}, got {}", expected, got)),
            crate::ooxml::error::OoxmlError::InvalidRelationship(msg) => XlsbError::Encoding(format!("Invalid relationship: {}", msg)),
            crate::ooxml::error::OoxmlError::InvalidFormat(msg) => XlsbError::Encoding(msg),
            crate::ooxml::error::OoxmlError::Io(e) => XlsbError::Io(e),
            crate::ooxml::error::OoxmlError::Other(msg) => XlsbError::Encoding(msg),
        }
    }
}

impl From<String> for XlsbError {
    fn from(err: String) -> Self {
        XlsbError::Encoding(err)
    }
}
