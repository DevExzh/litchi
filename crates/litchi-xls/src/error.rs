//! Error types for XLS file parsing

use std::fmt;

/// Result type alias for XLS operations
pub type Result<T, E = Error> = std::result::Result<T, E>;

/// Password-to-open encryption families identified in a BIFF `FILEPASS` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionKind {
    /// Legacy BIFF8 binary RC4.
    BinaryRc4,
    /// `CryptoAPI` encryption.
    CryptoApi,
    /// An unknown `wEncryptionType` value.
    Unknown(u16),
}

/// Resource dimension governed by an XLS source or edit limit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source bytes, before the CFB container is parsed.
    InputBytes,
    /// Bytes in one named CFB stream entry.
    EntryBytes,
    /// Aggregate bytes in all declared CFB stream entries.
    AggregateBytes,
    /// Changes retained by one primary cell edit.
    Operations,
    /// A dimension of the embedded OLE-object capture/publication policy.
    OleObject(crate::ole_object::LimitKind),
}

impl LimitKind {
    pub(crate) const fn description(self) -> &'static str {
        match self {
            Self::InputBytes => "input bytes",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
            Self::Operations => "staged changes",
            Self::OleObject(kind) => kind.description(),
        }
    }
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.description())
    }
}

/// Errors that can occur during XLS file parsing
#[non_exhaustive]
pub enum Error {
    /// I/O error
    Io(std::io::Error),
    /// CFB (Compound File Binary) error
    Cfb(litchi_cfb::OleError),
    /// Shared OLE object capture or publication failure.
    Object(litchi_ole_common::Error),
    /// MS-OVBA project authoring or parsing error
    Vba(litchi_vba::Error),
    /// Legacy Custom XML data-store parsing failure.
    CustomXml(litchi_ole_common::custom_xml::Error),
    /// XML signature verification failure.
    Signature(litchi_sign::Error),
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
    /// A well-formed edit request this route refuses, classified by
    /// [`litchi_core::edit::RefusalKind`]. The refusal is content-free: its
    /// operation and reason are static (record 0351 §2.4).
    Refused(litchi_core::edit::Refusal),
    /// A patch, plan or staged precondition was made against another source
    /// than the one receiving it.
    PatchConflict {
        /// The edit route that detected the conflict.
        operation: &'static str,
        /// Static explanation of the mismatch.
        reason: &'static str,
    },
    /// A changed write of a signed workbook (record 0351 §2.3 rule 5); remove
    /// the signature explicitly before editing.
    Signed,
    /// A broken internal invariant of an edit route (for example a readback
    /// that disagrees with the staged change); never caused by the request.
    Internal(&'static str),
    /// A caller supplied an invalid source or edit limit.
    InvalidLimit {
        /// Dimension whose configured value is invalid.
        kind: LimitKind,
        /// Requested value.
        value: u64,
        /// Hard safety ceiling.
        maximum: u64,
    },
    /// A configured source or edit budget was exceeded.
    Limit {
        /// Dimension that exceeded its selected limit.
        kind: LimitKind,
        /// Observed value.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A string is longer than the BIFF8 field that would store it.
    ///
    /// The writer returns this before it produces any output; it never
    /// truncates a string to make it fit.
    StringTooLong {
        /// The field that would store the string, such as `"shared string"`.
        field: &'static str,
        /// The string's length in UTF-16 code units.
        utf16_units: usize,
        /// The most UTF-16 code units the field can store.
        limit: usize,
    },
    /// A BIFF8 collection cannot hold one more entry.
    ///
    /// The writer returns this before it changes anything.
    TooMany {
        /// The collection that is full, such as `"custom number formats"`.
        collection: &'static str,
        /// The most entries BIFF8 (or litchi's reader) can store.
        limit: usize,
    },
    /// A record's payload is longer than a BIFF8 record can hold, and the
    /// record has no continuation.
    ///
    /// The writer returns this before it produces any output; it never
    /// truncates or wraps a record length.
    RecordTooLong {
        /// The record, such as `"XFExt"`.
        record: &'static str,
        /// The payload length in bytes.
        bytes: usize,
        /// The most payload bytes the record can hold.
        limit: usize,
    },
    /// A decoded worksheet cannot be edited through the create-only worksheet API.
    ///
    /// Use a source-checked transaction when one exists for the selected XLS
    /// feature, or author a new worksheet through [`crate::writer::Builder`].
    SourceBoundWorksheetMutation {
        /// The attempted create-only operation.
        operation: &'static str,
    },
    /// Weak legacy encryption authoring requires an explicit opt-in policy.
    WeakEncryptionRequiresExplicitPolicy,
    /// RC4 authoring requires caller-supplied one-use salt and verifier bytes.
    EncryptionEntropyRequired,
    /// The metadata given to [`crate::Builder::set_metadata`] cannot be
    /// stored in the OLE property sets.
    Metadata(litchi_ole_common::property_set::summary_properties::MetadataError),
    /// Document properties were combined with `CryptoAPI` RC4 encryption,
    /// whose encrypted summary stream ([MS-OFFCRYPTO] 2.3.5.4) is not
    /// produced; the builder refuses rather than store them in plaintext.
    EncryptedMetadata,
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

impl fmt::Debug for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Preserve the established diagnostics for the native error arms.
        // The foreign boundaries added by the unified API deliberately expose
        // only their variant names; their typed causes remain available from
        // `source()`.
        match self {
            Error::Io(error) => formatter.debug_tuple("Error::Io").field(error).finish(),
            Error::Cfb(error) => formatter.debug_tuple("Error::Cfb").field(error).finish(),
            Error::Object(error) => formatter.debug_tuple("Error::Object").field(error).finish(),
            Error::Vba(_) => formatter.write_str("Error::Vba(..)"),
            Error::CustomXml(_) => formatter.write_str("Error::CustomXml(..)"),
            Error::Signature(_) => formatter.write_str("Error::Signature(..)"),
            Error::Graph(error) => formatter.debug_tuple("Error::Graph").field(error).finish(),
            Error::InvalidRecord {
                record_type,
                message,
            } => formatter
                .debug_struct("Error::InvalidRecord")
                .field("record_type", record_type)
                .field("message", message)
                .finish(),
            Error::UnsupportedBiffVersion(version) => formatter
                .debug_tuple("Error::UnsupportedBiffVersion")
                .field(version)
                .finish(),
            Error::PasswordProtected => formatter.write_str("Error::PasswordProtected"),
            Error::PasswordRequired => formatter.write_str("Error::PasswordRequired"),
            Error::InvalidPassword => formatter.write_str("Error::InvalidPassword"),
            Error::UnsupportedEncryption(kind) => formatter
                .debug_tuple("Error::UnsupportedEncryption")
                .field(kind)
                .finish(),
            Error::MalformedFilePass(message) => formatter
                .debug_tuple("Error::MalformedFilePass")
                .field(message)
                .finish(),
            Error::InvalidLength { expected, found } => formatter
                .debug_struct("Error::InvalidLength")
                .field("expected", expected)
                .field("found", found)
                .finish(),
            Error::Allocation(context) => formatter
                .debug_tuple("Error::Allocation")
                .field(context)
                .finish(),
            Error::UnexpectedEndOfStream(context) => formatter
                .debug_tuple("Error::UnexpectedEndOfStream")
                .field(context)
                .finish(),
            Error::InvalidFormula(message) => formatter
                .debug_tuple("Error::InvalidFormula")
                .field(message)
                .finish(),
            Error::InvalidCellReference(reference) => formatter
                .debug_tuple("Error::InvalidCellReference")
                .field(reference)
                .finish(),
            Error::WorksheetNotFound(name) => formatter
                .debug_tuple("Error::WorksheetNotFound")
                .field(name)
                .finish(),
            Error::InvalidFormat(code) => formatter
                .debug_tuple("Error::InvalidFormat")
                .field(code)
                .finish(),
            Error::Encoding(message) => formatter
                .debug_tuple("Error::Encoding")
                .field(message)
                .finish(),
            Error::UnsupportedFeature(feature) => formatter
                .debug_tuple("Error::UnsupportedFeature")
                .field(feature)
                .finish(),
            Error::InvalidData(message) => formatter
                .debug_tuple("Error::InvalidData")
                .field(message)
                .finish(),
            Error::Refused(refusal) => formatter
                .debug_tuple("Error::Refused")
                .field(refusal)
                .finish(),
            Error::PatchConflict { operation, reason } => formatter
                .debug_struct("Error::PatchConflict")
                .field("operation", operation)
                .field("reason", reason)
                .finish(),
            Error::Signed => formatter.write_str("Error::Signed"),
            Error::Internal(reason) => formatter
                .debug_tuple("Error::Internal")
                .field(reason)
                .finish(),
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => formatter
                .debug_struct("Error::InvalidLimit")
                .field("kind", kind)
                .field("value", value)
                .field("maximum", maximum)
                .finish(),
            Error::Limit {
                kind,
                observed,
                maximum,
            } => formatter
                .debug_struct("Error::Limit")
                .field("kind", kind)
                .field("observed", observed)
                .field("maximum", maximum)
                .finish(),
            Error::StringTooLong {
                field,
                utf16_units,
                limit,
            } => formatter
                .debug_struct("Error::StringTooLong")
                .field("field", field)
                .field("utf16_units", utf16_units)
                .field("limit", limit)
                .finish(),
            Error::TooMany { collection, limit } => formatter
                .debug_struct("Error::TooMany")
                .field("collection", collection)
                .field("limit", limit)
                .finish(),
            Error::RecordTooLong {
                record,
                bytes,
                limit,
            } => formatter
                .debug_struct("Error::RecordTooLong")
                .field("record", record)
                .field("bytes", bytes)
                .field("limit", limit)
                .finish(),
            Error::SourceBoundWorksheetMutation { operation } => formatter
                .debug_struct("Error::SourceBoundWorksheetMutation")
                .field("operation", operation)
                .finish(),
            Error::WeakEncryptionRequiresExplicitPolicy => {
                formatter.write_str("Error::WeakEncryptionRequiresExplicitPolicy")
            },
            Error::EncryptionEntropyRequired => {
                formatter.write_str("Error::EncryptionEntropyRequired")
            },
            Error::Metadata(error) => formatter
                .debug_tuple("Error::Metadata")
                .field(error)
                .finish(),
            Error::EncryptedMetadata => formatter.write_str("Error::EncryptedMetadata"),
            Error::UnexpectedRecordType { expected, found } => formatter
                .debug_struct("Error::UnexpectedRecordType")
                .field("expected", expected)
                .field("found", found)
                .finish(),
            Error::Eof(context) => formatter.debug_tuple("Error::Eof").field(context).finish(),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::Io(e) => write!(f, "I/O error: {e}"),
            Error::Cfb(e) => write!(f, "CFB error: {e}"),
            Error::Object(e) => write!(f, "OLE object error: {e}"),
            Error::Vba(_) => f.write_str("XLS VBA operation failed"),
            Error::CustomXml(_) => f.write_str("XLS Custom XML operation failed"),
            Error::Signature(_) => f.write_str("XLS signature operation failed"),
            Error::Graph(e) => write!(f, "Office Graph error: {e}"),
            Error::InvalidRecord {
                record_type,
                message,
            } => {
                write!(f, "Invalid record 0x{record_type:04X}: {message}")
            },
            Error::UnsupportedBiffVersion(version) => {
                write!(f, "Unsupported BIFF version: {version}")
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
                write!(f, "Invalid length: expected {expected}, found {found}")
            },
            Error::Allocation(context) => {
                write!(f, "Allocation failed while {context}")
            },
            Error::UnexpectedEndOfStream(context) => {
                write!(f, "Unexpected end of stream: {context}")
            },
            Error::InvalidFormula(msg) => {
                write!(f, "Invalid formula: {msg}")
            },
            Error::InvalidCellReference(ref_str) => {
                write!(f, "Invalid cell reference: {ref_str}")
            },
            Error::WorksheetNotFound(name) => {
                write!(f, "Worksheet '{name}' not found")
            },
            Error::InvalidFormat(code) => {
                write!(f, "Invalid format code: {code}")
            },
            Error::Encoding(msg) => {
                write!(f, "Encoding error: {msg}")
            },
            Error::UnsupportedFeature(feature) => {
                write!(f, "Unsupported feature: {feature}")
            },
            Error::InvalidData(msg) => {
                write!(f, "Invalid data: {msg}")
            },
            Error::Refused(refusal) => fmt::Display::fmt(refusal, f),
            Error::PatchConflict { operation, reason } => {
                write!(f, "XLS {operation} patch conflict: {reason}")
            },
            Error::Internal(reason) => write!(f, "XLS edit invariant failed: {reason}"),
            Error::Signed => f.write_str("changed XLS writes require explicit signature removal"),
            Error::Limit {
                kind: LimitKind::Operations,
                observed,
                maximum,
            } => write!(
                f,
                "XLS staged change limit exceeded: {observed} > {maximum}"
            ),
            Error::Limit {
                kind: LimitKind::OleObject(kind),
                observed,
                maximum,
            } => write!(
                f,
                "XLS {} limit exceeded: {observed} > {maximum}",
                kind.description()
            ),
            Error::InvalidLimit {
                kind: LimitKind::OleObject(kind),
                value,
                maximum,
            } => write!(
                f,
                "invalid XLS {} limit: {value} (maximum {maximum})",
                kind.description()
            ),
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                f,
                "invalid XLS source limit for {kind}: {value} (maximum {maximum})"
            ),
            Error::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                f,
                "XLS source limit exceeded for {}: {observed} > {maximum}",
                kind.description()
            ),
            Error::StringTooLong {
                field,
                utf16_units,
                limit,
            } => write!(
                f,
                "{field} has {utf16_units} UTF-16 code units; BIFF8 stores at most {limit}"
            ),
            Error::TooMany { collection, limit } => {
                write!(f, "{collection} are full: BIFF8 stores at most {limit}")
            },
            Error::RecordTooLong {
                record,
                bytes,
                limit,
            } => write!(
                f,
                "{record} record needs {bytes} payload bytes; a BIFF8 record holds at most {limit}"
            ),
            Error::SourceBoundWorksheetMutation { operation } => write!(
                f,
                "source-bound worksheet mutation refused: {operation}; use a source-checked transaction or author a new worksheet"
            ),
            Error::WeakEncryptionRequiresExplicitPolicy => write!(
                f,
                "BIFF8 XOR obfuscation requires an explicit weak-encryption policy"
            ),
            Error::EncryptionEntropyRequired => {
                write!(f, "BIFF8 RC4 authoring requires one-use EncryptionEntropy")
            },
            Error::Metadata(error) => write!(f, "XLS metadata error: {error}"),
            Error::EncryptedMetadata => write!(
                f,
                "XLS document properties cannot be written with CryptoAPI RC4 encryption"
            ),
            Error::UnexpectedRecordType { expected, found } => {
                write!(
                    f,
                    "Unexpected record type: expected 0x{expected:04X}, found 0x{found:04X}"
                )
            },
            Error::Eof(context) => {
                write!(f, "End of file: {context}")
            },
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Error::Io(e) => Some(e),
            Error::Cfb(e) => Some(e),
            Error::Object(e) => Some(e),
            Error::Vba(e) => Some(e),
            Error::CustomXml(e) => Some(e),
            Error::Signature(e) => Some(e),
            Error::Graph(e) => Some(e),
            Error::Metadata(e) => Some(e),
            _ => None,
        }
    }
}

impl Error {
    /// A content-free refusal of `operation` for `reason`.
    #[must_use]
    pub const fn refused(
        kind: litchi_core::edit::RefusalKind,
        operation: &'static str,
        reason: &'static str,
    ) -> Self {
        Self::Refused(litchi_core::edit::Refusal::new(kind, operation, reason))
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

impl From<litchi_ole_common::property_set::summary_properties::MetadataError> for Error {
    fn from(err: litchi_ole_common::property_set::summary_properties::MetadataError) -> Self {
        Error::Metadata(err)
    }
}

impl From<litchi_vba::Error> for Error {
    fn from(err: litchi_vba::Error) -> Self {
        Error::Vba(err)
    }
}

impl From<litchi_ole_common::custom_xml::Error> for Error {
    fn from(err: litchi_ole_common::custom_xml::Error) -> Self {
        Error::CustomXml(err)
    }
}

impl From<litchi_sign::Error> for Error {
    fn from(err: litchi_sign::Error) -> Self {
        Error::Signature(err)
    }
}

impl From<litchi_ograph::Error> for Error {
    fn from(err: litchi_ograph::Error) -> Self {
        match err {
            litchi_ograph::Error::Refused(refusal) => Self::Refused(refusal),
            litchi_ograph::Error::PatchConflict { operation, reason } => {
                Self::PatchConflict { operation, reason }
            },
            err => Self::Graph(err),
        }
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

/// Classify an XLS error for the format-neutral sheet error the formula
/// engine reports. The message is static: no sheet name, reference or
/// record content crosses.
impl From<Error> for litchi_core::sheet::Error {
    #[cold]
    #[inline(never)]
    fn from(error: Error) -> Self {
        use litchi_core::edit::RefusalKind;
        let refused = |kind, reason| Self::refused(kind, "read", reason);
        match error {
            Error::Object(litchi_ole_common::Error::Cfb(litchi_cfb::OleError::SourceChanged {
                ..
            })) => Self::SourceChanged,
            Error::Object(litchi_ole_common::Error::Cfb(
                litchi_cfb::OleError::Io(_) | litchi_cfb::OleError::Committed { .. },
            )) => Self::Io,
            Error::Object(litchi_ole_common::Error::Cfb(litchi_cfb::OleError::Allocation {
                ..
            })) => Self::Internal("the XLS object allocation failed"),
            Error::Object(litchi_ole_common::Error::Cfb(litchi_cfb::OleError::StreamNotFound)) => {
                refused(RefusalKind::NotFound, "the XLS object stream was not found")
            },
            Error::Object(
                litchi_ole_common::Error::Limit { .. }
                | litchi_ole_common::Error::Cfb(litchi_cfb::OleError::LimitExceeded { .. }),
            ) => refused(
                RefusalKind::UnsupportedContent,
                "the XLS workbook exceeds a configured limit",
            ),
            Error::Object(
                litchi_ole_common::Error::InvalidLimit { .. }
                | litchi_ole_common::Error::Cfb(litchi_cfb::OleError::InvalidLimit { .. }),
            ) => refused(
                RefusalKind::InvalidValue,
                "the XLS object limit is not valid",
            ),
            Error::Object(litchi_ole_common::Error::PatchConflict) => Self::PatchConflict,
            Error::Io(_) | Error::Cfb(_) => Self::Io,
            Error::WorksheetNotFound(_) => refused(
                RefusalKind::NotFound,
                "the XLS workbook has no such worksheet",
            ),
            Error::InvalidCellReference(_) => {
                refused(RefusalKind::InvalidValue, "the cell reference is not valid")
            },
            Error::PasswordProtected
            | Error::PasswordRequired
            | Error::InvalidPassword
            | Error::UnsupportedEncryption(_) => {
                refused(RefusalKind::Protected, "the XLS workbook is encrypted")
            },
            Error::UnsupportedBiffVersion(_) | Error::UnsupportedFeature(_) => refused(
                RefusalKind::UnsupportedContent,
                "the XLS workbook uses an unsupported feature",
            ),
            Error::Limit { .. } | Error::InvalidLimit { .. } | Error::Allocation(_) => refused(
                RefusalKind::UnsupportedContent,
                "the XLS workbook exceeds a configured limit",
            ),
            _ => Self::Malformed("the XLS data is malformed"),
        }
    }
}

impl From<litchi_ole_common::Error> for Error {
    fn from(error: litchi_ole_common::Error) -> Self {
        Self::Object(error)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn object_errors_keep_sheet_failure_categories() {
        use litchi_core::{edit::RefusalKind, sheet::Error as SheetError};
        use litchi_ole_common::{Error as ObjectError, LimitKind};
        let convert = |error| SheetError::from(Error::Object(error));
        assert_eq!(
            convert(ObjectError::Cfb(litchi_cfb::OleError::Io(
                std::io::Error::other("inert")
            ))),
            SheetError::Io
        );
        assert_eq!(
            convert(ObjectError::Cfb(litchi_cfb::OleError::SourceChanged {
                expected: litchi_core::SourceVersion::new(1, 0),
                observed: litchi_core::SourceVersion::new(1, 1),
            })),
            SheetError::SourceChanged
        );
        assert_eq!(
            convert(ObjectError::PatchConflict),
            SheetError::PatchConflict
        );
        assert_eq!(
            convert(ObjectError::Limit {
                kind: LimitKind::StreamBytes,
                observed: 2,
                maximum: 1
            })
            .refusal()
            .map(litchi_core::edit::Refusal::kind),
            Some(RefusalKind::UnsupportedContent)
        );
        assert_eq!(
            convert(ObjectError::InvalidLimit {
                kind: LimitKind::StorageDepth,
                value: 0,
                maximum: 32
            })
            .refusal()
            .map(litchi_core::edit::Refusal::kind),
            Some(RefusalKind::InvalidValue)
        );
        assert_eq!(
            convert(ObjectError::Cfb(litchi_cfb::OleError::LimitExceeded {
                resource: "input bytes",
                observed: 2,
                maximum: 1,
            }))
            .refusal()
            .map(litchi_core::edit::Refusal::kind),
            Some(RefusalKind::UnsupportedContent)
        );
        assert_eq!(
            convert(ObjectError::Cfb(litchi_cfb::OleError::InvalidLimit {
                resource: "input bytes",
                value: 0,
                maximum: 1,
            }))
            .refusal()
            .map(litchi_core::edit::Refusal::kind),
            Some(RefusalKind::InvalidValue)
        );
        assert!(matches!(
            convert(ObjectError::Cfb(litchi_cfb::OleError::InvalidFormat(
                "inert".into()
            ))),
            SheetError::Malformed(_)
        ));
    }
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
    fn test_xls_error_string_too_long() {
        let err = Error::StringTooLong {
            field: "shared string",
            utf16_units: 65_536,
            limit: 65_535,
        };
        assert_eq!(
            err.to_string(),
            "shared string has 65536 UTF-16 code units; BIFF8 stores at most 65535"
        );
        assert!(err.source().is_none());
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

    #[test]
    fn new_foreign_boundaries_redact_payloads_but_keep_typed_sources() {
        let payload = "xls-secret-foreign-payload";
        let errors = [
            Error::Vba(litchi_vba::Error::InvalidData(payload.to_string())),
            Error::CustomXml(litchi_ole_common::custom_xml::Error::Invalid(
                payload.to_string(),
            )),
            Error::Signature(litchi_sign::Error::Xml(payload.to_string())),
        ];
        for error in errors {
            assert!(error.source().is_some());
            assert!(!error.to_string().contains(payload));
            assert!(!format!("{error:?}").contains(payload));
        }
    }
}
