//! Error types for XLS file parsing

use std::fmt;

/// Result type alias for XLS operations
pub type XlsResult<T> = Result<T, XlsError>;

/// Errors that can occur during XLS file parsing
#[derive(Debug)]
pub enum XlsError {
    /// I/O error
    Io(std::io::Error),
    /// CFB (Compound File Binary) error
    Cfb(crate::ole::file::OleError),
    /// Invalid BIFF record
    InvalidRecord {
        /// Record type
        record_type: u16,
        /// Error description
        message: String,
    },
    /// Unsupported BIFF version
    UnsupportedBiffVersion(u16),
    /// Password protected workbook
    PasswordProtected,
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
    /// Invalid format code
    InvalidFormat(u16),
    /// Encoding error
    Encoding(String),
    /// Unsupported feature
    UnsupportedFeature(String),
    /// Invalid data
    InvalidData(String),
    /// Unexpected record type
    UnexpectedRecordType {
        /// Expected record type
        expected: u16,
        /// Found record type
        found: u16,
    },
    /// End of file
    Eof(&'static str),
}

impl fmt::Display for XlsError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            XlsError::Io(e) => write!(f, "I/O error: {}", e),
            XlsError::Cfb(e) => write!(f, "CFB error: {}", e),
            XlsError::InvalidRecord {
                record_type,
                message,
            } => {
                write!(f, "Invalid record 0x{:04X}: {}", record_type, message)
            },
            XlsError::UnsupportedBiffVersion(version) => {
                write!(f, "Unsupported BIFF version: {}", version)
            },
            XlsError::PasswordProtected => {
                write!(f, "Workbook is password protected")
            },
            XlsError::InvalidLength { expected, found } => {
                write!(f, "Invalid length: expected {}, found {}", expected, found)
            },
            XlsError::UnexpectedEndOfStream(context) => {
                write!(f, "Unexpected end of stream: {}", context)
            },
            XlsError::InvalidFormula(msg) => {
                write!(f, "Invalid formula: {}", msg)
            },
            XlsError::InvalidCellReference(ref_str) => {
                write!(f, "Invalid cell reference: {}", ref_str)
            },
            XlsError::WorksheetNotFound(name) => {
                write!(f, "Worksheet '{}' not found", name)
            },
            XlsError::InvalidFormat(code) => {
                write!(f, "Invalid format code: {}", code)
            },
            XlsError::Encoding(msg) => {
                write!(f, "Encoding error: {}", msg)
            },
            XlsError::UnsupportedFeature(feature) => {
                write!(f, "Unsupported feature: {}", feature)
            },
            XlsError::InvalidData(msg) => {
                write!(f, "Invalid data: {}", msg)
            },
            XlsError::UnexpectedRecordType { expected, found } => {
                write!(
                    f,
                    "Unexpected record type: expected 0x{:04X}, found 0x{:04X}",
                    expected, found
                )
            },
            XlsError::Eof(context) => {
                write!(f, "End of file: {}", context)
            },
        }
    }
}

impl std::error::Error for XlsError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            XlsError::Io(e) => Some(e),
            XlsError::Cfb(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for XlsError {
    fn from(err: std::io::Error) -> Self {
        XlsError::Io(err)
    }
}

impl From<crate::ole::file::OleError> for XlsError {
    fn from(err: crate::ole::file::OleError) -> Self {
        XlsError::Cfb(err)
    }
}

impl From<crate::common::binary::BinaryError> for XlsError {
    fn from(err: crate::common::binary::BinaryError) -> Self {
        XlsError::InvalidData(err.to_string())
    }
}
