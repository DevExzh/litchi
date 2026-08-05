//! Error types for XLS file parsing

use std::fmt;

/// Result type alias for XLS operations
pub type XlsResult<T> = Result<T, XlsError>;

/// Password-to-open encryption families identified in a BIFF `FILEPASS` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsEncryptionKind {
    /// Legacy BIFF8 binary RC4.
    BinaryRc4,
    /// CryptoAPI encryption.
    CryptoApi,
    /// An unknown `wEncryptionType` value.
    Unknown(u16),
}

/// Errors that can occur during XLS file parsing
#[derive(Debug)]
pub enum XlsError {
    /// I/O error
    Io(std::io::Error),
    /// CFB (Compound File Binary) error
    Cfb(litchi_cfb::OleError),
    /// MS-OVBA project authoring or parsing error
    Vba(litchi_vba::Error),
    /// Host-neutral Office Graph parsing or encoding error.
    Graph(litchi_ograph::Error),
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
    /// A password-to-open encrypted workbook was opened without a password.
    PasswordRequired,
    /// The supplied password did not match the `FILEPASS` verifier.
    InvalidPassword,
    /// The encryption family is recognized but not implemented.
    UnsupportedEncryption(XlsEncryptionKind),
    /// The `FILEPASS` record is structurally invalid.
    MalformedFilePass(String),
    /// Invalid data length
    InvalidLength {
        /// Expected length
        expected: usize,
        /// Found length
        found: usize,
    },
    /// A fallible retained-state or parsing allocation could not be reserved.
    Allocation(&'static str),
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
    /// A requested edit cannot preserve unsupported source records safely.
    UnsafeEdit(String),
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
            XlsError::Vba(e) => write!(f, "VBA project error: {}", e),
            XlsError::Graph(e) => write!(f, "Office Graph error: {e}"),
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
            XlsError::PasswordRequired => write!(f, "Workbook password is required"),
            XlsError::InvalidPassword => write!(f, "Invalid workbook password"),
            XlsError::UnsupportedEncryption(kind) => {
                write!(f, "Unsupported workbook encryption: {kind:?}")
            },
            XlsError::MalformedFilePass(message) => {
                write!(f, "Malformed FILEPASS record: {message}")
            },
            XlsError::InvalidLength { expected, found } => {
                write!(f, "Invalid length: expected {}, found {}", expected, found)
            },
            XlsError::Allocation(context) => {
                write!(f, "Allocation failed while {context}")
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
            XlsError::UnsafeEdit(msg) => write!(f, "Unsafe edit refused: {msg}"),
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
            XlsError::Vba(e) => Some(e),
            XlsError::Graph(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for XlsError {
    fn from(err: std::io::Error) -> Self {
        XlsError::Io(err)
    }
}

impl From<litchi_cfb::OleError> for XlsError {
    fn from(err: litchi_cfb::OleError) -> Self {
        XlsError::Cfb(err)
    }
}

impl From<litchi_vba::Error> for XlsError {
    fn from(err: litchi_vba::Error) -> Self {
        XlsError::Vba(err)
    }
}

impl From<litchi_ograph::Error> for XlsError {
    fn from(err: litchi_ograph::Error) -> Self {
        Self::Graph(err)
    }
}

impl From<litchi_core::binary::BinaryError> for XlsError {
    fn from(err: litchi_core::binary::BinaryError) -> Self {
        XlsError::InvalidData(err.to_string())
    }
}

impl From<litchi_biff::Error> for XlsError {
    fn from(err: litchi_biff::Error) -> Self {
        match err {
            litchi_biff::Error::TruncatedHeader { available, .. } => XlsError::InvalidLength {
                expected: 4,
                found: available,
            },
            litchi_biff::Error::TruncatedPayload {
                kind,
                declared,
                available,
                ..
            } => XlsError::InvalidRecord {
                record_type: kind.get(),
                message: format!(
                    "truncated BIFF payload: declared {declared} bytes, only {available} available"
                ),
            },
            other => XlsError::InvalidData(other.to_string()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::error::Error;

    #[test]
    fn test_xls_error_io() {
        let io_err = std::io::Error::new(std::io::ErrorKind::NotFound, "file not found");
        let err = XlsError::Io(io_err);
        let display = format!("{}", err);
        assert!(display.contains("I/O error"));
        assert!(display.contains("file not found"));
    }

    #[test]
    fn test_xls_error_invalid_record() {
        let err = XlsError::InvalidRecord {
            record_type: 0x0201,
            message: "Invalid data".to_string(),
        };
        let display = format!("{}", err);
        assert!(display.contains("Invalid record 0x0201"));
        assert!(display.contains("Invalid data"));
    }

    #[test]
    fn test_xls_error_unsupported_biff_version() {
        let err = XlsError::UnsupportedBiffVersion(0x0100);
        let display = format!("{}", err);
        assert!(display.contains("Unsupported BIFF version: 256"));
    }

    #[test]
    fn test_xls_error_password_protected() {
        let err = XlsError::PasswordProtected;
        let display = format!("{}", err);
        assert_eq!(display, "Workbook is password protected");
    }

    #[test]
    fn test_xls_error_invalid_length() {
        let err = XlsError::InvalidLength {
            expected: 10,
            found: 5,
        };
        let display = format!("{}", err);
        assert_eq!(display, "Invalid length: expected 10, found 5");
    }

    #[test]
    fn test_xls_error_unexpected_end_of_stream() {
        let err = XlsError::UnexpectedEndOfStream("while reading header".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Unexpected end of stream"));
        assert!(display.contains("while reading header"));
    }

    #[test]
    fn test_xls_error_invalid_formula() {
        let err = XlsError::InvalidFormula("Missing closing paren".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Invalid formula"));
        assert!(display.contains("Missing closing paren"));
    }

    #[test]
    fn test_xls_error_invalid_cell_reference() {
        let err = XlsError::InvalidCellReference("XYZ123".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Invalid cell reference: XYZ123"));
    }

    #[test]
    fn test_xls_error_worksheet_not_found() {
        let err = XlsError::WorksheetNotFound("Sheet99".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Worksheet 'Sheet99' not found"));
    }

    #[test]
    fn test_xls_error_invalid_format() {
        let err = XlsError::InvalidFormat(0xFF);
        let display = format!("{}", err);
        assert!(display.contains("Invalid format code: 255"));
    }

    #[test]
    fn test_xls_error_encoding() {
        let err = XlsError::Encoding("UTF-8 error".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Encoding error"));
        assert!(display.contains("UTF-8 error"));
    }

    #[test]
    fn test_xls_error_unsupported_feature() {
        let err = XlsError::UnsupportedFeature("Pivot tables".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Unsupported feature"));
        assert!(display.contains("Pivot tables"));
    }

    #[test]
    fn test_xls_error_invalid_data() {
        let err = XlsError::InvalidData("Corrupted header".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Invalid data"));
        assert!(display.contains("Corrupted header"));
    }

    #[test]
    fn test_xls_error_unexpected_record_type() {
        let err = XlsError::UnexpectedRecordType {
            expected: 0x0009,
            found: 0x0006,
        };
        let display = format!("{}", err);
        assert!(display.contains("Unexpected record type"));
        assert!(display.contains("expected 0x0009"));
        assert!(display.contains("found 0x0006"));
    }

    #[test]
    fn test_xls_error_eof() {
        let err = XlsError::Eof("stream");
        let display = format!("{}", err);
        assert!(display.contains("End of file: stream"));
    }

    #[test]
    fn test_xls_error_debug() {
        let err = XlsError::PasswordProtected;
        let debug_str = format!("{:?}", err);
        assert!(debug_str.contains("PasswordProtected"));
    }

    #[test]
    fn test_xls_error_source_io() {
        let io_err = std::io::Error::new(std::io::ErrorKind::NotFound, "test");
        let err = XlsError::Io(io_err);
        assert!(err.source().is_some());
    }

    #[test]
    fn test_xls_error_source_other() {
        let err = XlsError::PasswordProtected;
        assert!(err.source().is_none());
    }

    #[test]
    fn test_xls_error_from_io() {
        let io_err = std::io::Error::new(std::io::ErrorKind::PermissionDenied, "access denied");
        let xls_err: XlsError = io_err.into();
        match xls_err {
            XlsError::Io(_) => {},
            _ => panic!("Expected Io error variant"),
        }
    }

    #[test]
    fn test_xls_result_type() {
        fn returns_ok() -> XlsResult<u32> {
            Ok(42)
        }
        fn returns_err() -> XlsResult<u32> {
            Err(XlsError::PasswordProtected)
        }

        assert_eq!(returns_ok().unwrap(), 42);
        assert!(returns_err().is_err());
    }

    #[test]
    fn test_xls_error_cfb_source() {
        // Test that CFB errors properly return source
        let cfb_err = litchi_cfb::OleError::StreamNotFound;
        let err = XlsError::Cfb(cfb_err);
        // The CFB error doesn't have a source (unit variant), but test the variant
        let display = format!("{}", err);
        assert!(display.contains("Stream not found"));
    }
}
