//! Error types for XLS file parsing

use std::fmt;

/// Result type alias for XLS operations
pub type Result<T, E = Error> = std::result::Result<T, E>;

/// Password-to-open encryption families identified in a BIFF `FILEPASS` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionKind {
    /// Legacy BIFF8 binary RC4.
    BinaryRc4,
    /// CryptoAPI encryption.
    CryptoApi,
    /// An unknown `wEncryptionType` value.
    Unknown(u16),
}

/// Errors that can occur during XLS file parsing
#[derive(Debug)]
pub enum Error {
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
    UnsupportedEncryption(EncryptionKind),
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
    /// A decoded worksheet cannot be edited through the create-only worksheet API.
    ///
    /// Use a source-checked transaction when one exists for the selected XLS
    /// feature, or author a new worksheet through [`crate::writer::Writer`].
    SourceBoundWorksheetMutation {
        /// The attempted create-only operation.
        operation: &'static str,
    },
    /// Weak legacy encryption authoring requires an explicit opt-in policy.
    WeakEncryptionRequiresExplicitPolicy,
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

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::Io(e) => write!(f, "I/O error: {}", e),
            Error::Cfb(e) => write!(f, "CFB error: {}", e),
            Error::Vba(e) => write!(f, "VBA project error: {}", e),
            Error::Graph(e) => write!(f, "Office Graph error: {e}"),
            Error::InvalidRecord {
                record_type,
                message,
            } => {
                write!(f, "Invalid record 0x{:04X}: {}", record_type, message)
            },
            Error::UnsupportedBiffVersion(version) => {
                write!(f, "Unsupported BIFF version: {}", version)
            },
            Error::PasswordProtected => {
                write!(f, "Workbook is password protected")
            },
            Error::PasswordRequired => write!(f, "Workbook password is required"),
            Error::InvalidPassword => write!(f, "Invalid workbook password"),
            Error::UnsupportedEncryption(kind) => {
                write!(f, "Unsupported workbook encryption: {kind:?}")
            },
            Error::MalformedFilePass(message) => {
                write!(f, "Malformed FILEPASS record: {message}")
            },
            Error::InvalidLength { expected, found } => {
                write!(f, "Invalid length: expected {}, found {}", expected, found)
            },
            Error::Allocation(context) => {
                write!(f, "Allocation failed while {context}")
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
            Error::InvalidFormat(code) => {
                write!(f, "Invalid format code: {}", code)
            },
            Error::Encoding(msg) => {
                write!(f, "Encoding error: {}", msg)
            },
            Error::UnsupportedFeature(feature) => {
                write!(f, "Unsupported feature: {}", feature)
            },
            Error::InvalidData(msg) => {
                write!(f, "Invalid data: {}", msg)
            },
            Error::UnsafeEdit(msg) => write!(f, "Unsafe edit refused: {msg}"),
            Error::SourceBoundWorksheetMutation { operation } => write!(
                f,
                "source-bound worksheet mutation refused: {operation}; use a source-checked transaction or author a new worksheet"
            ),
            Error::WeakEncryptionRequiresExplicitPolicy => write!(
                f,
                "BIFF8 XOR obfuscation requires an explicit weak-encryption policy"
            ),
            Error::UnexpectedRecordType { expected, found } => {
                write!(
                    f,
                    "Unexpected record type: expected 0x{:04X}, found 0x{:04X}",
                    expected, found
                )
            },
            Error::Eof(context) => {
                write!(f, "End of file: {}", context)
            },
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Error::Io(e) => Some(e),
            Error::Cfb(e) => Some(e),
            Error::Vba(e) => Some(e),
            Error::Graph(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for Error {
    fn from(err: std::io::Error) -> Self {
        Error::Io(err)
    }
}

impl From<litchi_cfb::OleError> for Error {
    fn from(err: litchi_cfb::OleError) -> Self {
        Error::Cfb(err)
    }
}

impl From<litchi_vba::Error> for Error {
    fn from(err: litchi_vba::Error) -> Self {
        Error::Vba(err)
    }
}

impl From<litchi_ograph::Error> for Error {
    fn from(err: litchi_ograph::Error) -> Self {
        Self::Graph(err)
    }
}

impl From<litchi_core::binary::BinaryError> for Error {
    fn from(err: litchi_core::binary::BinaryError) -> Self {
        Error::InvalidData(err.to_string())
    }
}

impl From<litchi_biff::Error> for Error {
    fn from(err: litchi_biff::Error) -> Self {
        match err {
            litchi_biff::Error::TruncatedHeader { available, .. } => Error::InvalidLength {
                expected: 4,
                found: available,
            },
            litchi_biff::Error::TruncatedPayload {
                kind,
                declared,
                available,
                ..
            } => Error::InvalidRecord {
                record_type: kind.get(),
                message: format!(
                    "truncated BIFF payload: declared {declared} bytes, only {available} available"
                ),
            },
            other => Error::InvalidData(other.to_string()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::error::Error as StdError;

    #[test]
    fn test_xls_error_io() {
        let io_err = std::io::Error::new(std::io::ErrorKind::NotFound, "file not found");
        let err = Error::Io(io_err);
        let display = format!("{}", err);
        assert!(display.contains("I/O error"));
        assert!(display.contains("file not found"));
    }

    #[test]
    fn test_xls_error_invalid_record() {
        let err = Error::InvalidRecord {
            record_type: 0x0201,
            message: "Invalid data".to_string(),
        };
        let display = format!("{}", err);
        assert!(display.contains("Invalid record 0x0201"));
        assert!(display.contains("Invalid data"));
    }

    #[test]
    fn test_xls_error_unsupported_biff_version() {
        let err = Error::UnsupportedBiffVersion(0x0100);
        let display = format!("{}", err);
        assert!(display.contains("Unsupported BIFF version: 256"));
    }

    #[test]
    fn test_xls_error_password_protected() {
        let err = Error::PasswordProtected;
        let display = format!("{}", err);
        assert_eq!(display, "Workbook is password protected");
    }

    #[test]
    fn test_xls_error_invalid_length() {
        let err = Error::InvalidLength {
            expected: 10,
            found: 5,
        };
        let display = format!("{}", err);
        assert_eq!(display, "Invalid length: expected 10, found 5");
    }

    #[test]
    fn test_xls_error_unexpected_end_of_stream() {
        let err = Error::UnexpectedEndOfStream("while reading header".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Unexpected end of stream"));
        assert!(display.contains("while reading header"));
    }

    #[test]
    fn test_xls_error_invalid_formula() {
        let err = Error::InvalidFormula("Missing closing paren".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Invalid formula"));
        assert!(display.contains("Missing closing paren"));
    }

    #[test]
    fn test_xls_error_invalid_cell_reference() {
        let err = Error::InvalidCellReference("XYZ123".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Invalid cell reference: XYZ123"));
    }

    #[test]
    fn test_xls_error_worksheet_not_found() {
        let err = Error::WorksheetNotFound("Sheet99".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Worksheet 'Sheet99' not found"));
    }

    #[test]
    fn test_xls_error_invalid_format() {
        let err = Error::InvalidFormat(0xFF);
        let display = format!("{}", err);
        assert!(display.contains("Invalid format code: 255"));
    }

    #[test]
    fn test_xls_error_encoding() {
        let err = Error::Encoding("UTF-8 error".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Encoding error"));
        assert!(display.contains("UTF-8 error"));
    }

    #[test]
    fn test_xls_error_unsupported_feature() {
        let err = Error::UnsupportedFeature("Pivot tables".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Unsupported feature"));
        assert!(display.contains("Pivot tables"));
    }

    #[test]
    fn test_xls_error_invalid_data() {
        let err = Error::InvalidData("Corrupted header".to_string());
        let display = format!("{}", err);
        assert!(display.contains("Invalid data"));
        assert!(display.contains("Corrupted header"));
    }

    #[test]
    fn test_xls_error_unexpected_record_type() {
        let err = Error::UnexpectedRecordType {
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
        let err = Error::Eof("stream");
        let display = format!("{}", err);
        assert!(display.contains("End of file: stream"));
    }

    #[test]
    fn test_xls_error_debug() {
        let err = Error::PasswordProtected;
        let debug_str = format!("{:?}", err);
        assert!(debug_str.contains("PasswordProtected"));
    }

    #[test]
    fn test_xls_error_source_io() {
        let io_err = std::io::Error::new(std::io::ErrorKind::NotFound, "test");
        let err = Error::Io(io_err);
        assert!(err.source().is_some());
    }

    #[test]
    fn test_xls_error_source_other() {
        let err = Error::PasswordProtected;
        assert!(err.source().is_none());
    }

    #[test]
    fn test_xls_error_from_io() {
        let io_err = std::io::Error::new(std::io::ErrorKind::PermissionDenied, "access denied");
        let xls_err: Error = io_err.into();
        match xls_err {
            Error::Io(_) => {},
            _ => panic!("Expected Io error variant"),
        }
    }

    #[test]
    fn test_xls_result_type() {
        fn returns_ok() -> Result<u32> {
            Ok(42)
        }
        fn returns_err() -> Result<u32> {
            Err(Error::PasswordProtected)
        }

        assert_eq!(returns_ok().unwrap(), 42);
        assert!(returns_err().is_err());
    }

    #[test]
    fn test_xls_error_cfb_source() {
        // Test that CFB errors properly return source
        let cfb_err = litchi_cfb::OleError::StreamNotFound;
        let err = Error::Cfb(cfb_err);
        // The CFB error doesn't have a source (unit variant), but test the variant
        let display = format!("{}", err);
        assert!(display.contains("Stream not found"));
    }
}
