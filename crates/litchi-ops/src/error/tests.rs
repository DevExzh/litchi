//! One test per row of the record 0003 §2.6 classification table (with the
//! record 0007 Q1 codes), one per reachable wildcard arm, the size assertion,
//! and the `Display`/suggestion format.

use std::collections::HashSet;
use std::collections::TryReserveError;
use std::io;
use std::num::NonZeroUsize;
use std::sync::Arc;

use litchi_core::{ExecutionError, FileFormat, Resource, ResourceLimit, SourceVersion};

#[cfg(any(feature = "doc", feature = "xls", feature = "ppt", feature = "xlsb"))]
use super::PASSWORD_SUGGESTION;
use super::{Code, Error, LIMIT_SUGGESTION, SIGNED_SUGGESTION, Source};

fn allocation_failure() -> TryReserveError {
    Vec::<u8>::new()
        .try_reserve_exact(usize::MAX)
        .expect_err("reserving usize::MAX bytes must fail without allocating")
}

fn resource_limit() -> ResourceLimit {
    ResourceLimit {
        resource: Resource::InputBytes,
        observed: 11,
        limit: 10,
        scope: Arc::from("test input"),
    }
}

fn changed() -> (SourceVersion, SourceVersion) {
    (SourceVersion::new(1, 1), SourceVersion::new(1, 2))
}

#[track_caller]
fn assert_code(error: impl Into<Error>, code: Code) -> Error {
    let error = error.into();
    assert_eq!(error.code(), code, "{error:?}");
    error
}

#[track_caller]
fn assert_suggested(error: impl Into<Error>, code: Code, suggestion: &str) {
    let error = assert_code(error, code);
    assert_eq!(error.suggestion(), Some(suggestion), "{error:?}");
}

#[test]
fn error_is_one_pointer_wide() {
    assert_eq!(size_of::<Error>(), size_of::<usize>());
    assert_eq!(size_of::<super::Result<()>>(), size_of::<usize>());
}

#[test]
fn error_is_send_sync_and_static() {
    fn assert_bounds<T: Send + Sync + 'static>() {}
    assert_bounds::<Error>();
}

#[test]
fn codes_are_a_closed_list_of_unique_snake_case_wire_strings() {
    // 36 OfficeCLI codes plus `unsupported_capability`, `limit_exceeded` and
    // `cancelled`.
    assert_eq!(Code::ALL.len(), 39);
    let spellings: HashSet<&str> = Code::ALL.iter().map(|code| code.as_str()).collect();
    assert_eq!(spellings.len(), Code::ALL.len());
    for code in Code::ALL {
        let spelling = code.as_str();
        assert!(
            spelling
                .bytes()
                .all(|byte| byte.is_ascii_lowercase() || byte == b'_'),
            "{spelling}"
        );
        assert_eq!(code.to_string(), spelling);
    }
    assert_eq!(
        Code::UnsupportedCapability.as_str(),
        "unsupported_capability"
    );
    assert_eq!(Code::LimitExceeded.as_str(), "limit_exceeded");
    assert_eq!(Code::Cancelled.as_str(), "cancelled");
    assert_eq!(Code::InvalidXpath.as_str(), "invalid_xpath");
}

#[test]
fn display_renders_code_and_message() {
    let error = Error::new(Code::InvalidPath, "path '/body/p[0]' is not one-based");
    assert_eq!(
        error.to_string(),
        "invalid_path: path '/body/p[0]' is not one-based"
    );
    assert_eq!(error.message(), "path '/body/p[0]' is not one-based");
    assert_eq!(error.suggestion(), None);
    assert!(matches!(error.leaf(), Source::None));
    assert!(std::error::Error::source(&error).is_none());
}

#[test]
fn display_appends_the_suggestion() {
    let error = Error::new(Code::UnsupportedProperty, "unknown property 'font.bolld'")
        .with_suggestion("did you mean 'font.bold'?")
        .with_format(FileFormat::Docx);
    assert_eq!(
        error.to_string(),
        "unsupported_property: unknown property 'font.bolld' (suggestion: did you mean 'font.bold'?)"
    );
    assert_eq!(error.suggestion(), Some("did you mean 'font.bold'?"));
    assert_eq!(error.format(), Some(FileFormat::Docx));
}

#[test]
fn source_returns_the_typed_leaf_error() {
    let error = Error::from(litchi_core::Error::ResourceLimit(resource_limit()));
    let source = std::error::Error::source(&error).expect("leaf source");
    let leaf = source
        .downcast_ref::<litchi_core::Error>()
        .expect("typed core error");
    let litchi_core::Error::ResourceLimit(limit) = leaf else {
        panic!("expected a resource limit, got {leaf:?}");
    };
    assert_eq!((limit.observed, limit.limit), (11, 10));
    assert_eq!(error.message(), leaf.to_string());
    assert!(matches!(error.leaf(), Source::Core(_)));
}

// S2 follow-up: `Display` is facade-level, so an error-chain printer that
// walks `source()` shows the leaf's text once, as the cause.
#[test]
fn display_does_not_repeat_the_leaf_text() {
    let error = Error::from(litchi_core::Error::ResourceLimit(resource_limit()));
    let leaf = std::error::Error::source(&error)
        .expect("leaf source")
        .to_string();
    let rendered = error.to_string();
    assert_eq!(
        rendered,
        "limit_exceeded: a resource limit was exceeded (suggestion: raise the matching limit in the open options if the input is trusted)"
    );
    assert!(!rendered.contains(&leaf), "{rendered:?} repeats {leaf:?}");
    assert_eq!(error.message(), leaf);

    #[cfg(feature = "docx")]
    {
        let error = Error::from(litchi_docx::Error::InvalidFormat("bad body".into()));
        assert_eq!(
            error.to_string(),
            "corrupt_file: the file is malformed [Docx]"
        );
        assert!(error.message().contains("bad body"));
    }
}

#[test]
fn every_code_has_a_summary() {
    let summaries: HashSet<&str> = Code::ALL.iter().map(|code| code.summary()).collect();
    assert_eq!(summaries.len(), Code::ALL.len(), "summaries are distinct");
    for code in Code::ALL {
        assert!(!code.summary().is_empty(), "{code}");
    }
}

#[test]
fn clones_share_the_leaf_error() {
    let error = Error::from(io::Error::other("disk on fire"));
    let clone = error.clone().with_suggestion("try another disk");
    assert_eq!(clone.code(), Code::IoError);
    assert_eq!(clone.suggestion(), Some("try another disk"));
    assert_eq!(error.suggestion(), None, "the original is unchanged");
    assert!(
        std::ptr::eq(error.leaf(), clone.leaf()),
        "the leaf is shared"
    );
}

// Row: `io::ErrorKind::NotFound` -> `file_not_found`.
#[test]
fn row_io_not_found_is_file_not_found() {
    let not_found = || io::Error::from(io::ErrorKind::NotFound);
    assert_code(not_found(), Code::FileNotFound);
    assert_code(litchi_core::Error::Io(not_found()), Code::FileNotFound);
    #[cfg(feature = "docx")]
    {
        assert_code(
            litchi_opc::OpcError::IoError(not_found()),
            Code::FileNotFound,
        );
        assert_code(
            litchi_opc::OpcError::PackageNotFound("missing.docx".into()),
            Code::FileNotFound,
        );
        let error = assert_code(litchi_docx::Error::Io(not_found()), Code::FileNotFound);
        assert_eq!(error.format(), Some(FileFormat::Docx));
    }
    #[cfg(feature = "pptx")]
    assert_code(litchi_pptx::Error::Io(not_found()), Code::FileNotFound);
    #[cfg(feature = "xlsb")]
    assert_code(
        litchi_xlsb::package::PackageError::Io(not_found()),
        Code::FileNotFound,
    );
    #[cfg(feature = "doc")]
    {
        assert_code(litchi_cfb::OleError::Io(not_found()), Code::FileNotFound);
        assert_code(litchi_doc::Error::Io(not_found()), Code::FileNotFound);
    }
    #[cfg(feature = "xls")]
    assert_code(litchi_xls::Error::Io(not_found()), Code::FileNotFound);
    #[cfg(feature = "ppt")]
    assert_code(litchi_ppt::Error::Io(not_found()), Code::FileNotFound);
}

// Row: other `io::Error`, `Error::SourceChanged`, OPC `SourceChanged` ->
// `io_error`; `SourceChanged` keeps its own message.
#[test]
fn row_other_io_and_source_changes_are_io_error() {
    assert_code(io::Error::other("disk on fire"), Code::IoError);
    let (expected, observed) = changed();
    let leaf = litchi_core::Error::SourceChanged { expected, observed };
    let rendered = leaf.to_string();
    let error = assert_code(leaf, Code::IoError);
    assert_eq!(error.message(), rendered);
    #[cfg(feature = "docx")]
    {
        let (expected, actual) = changed();
        assert_code(
            litchi_opc::OpcError::SourceChanged { expected, actual },
            Code::IoError,
        );
        let (expected, actual) = changed();
        assert_code(
            litchi_docx::Error::Opc(litchi_opc::OpcError::SourceChanged { expected, actual }),
            Code::IoError,
        );
    }
    #[cfg(feature = "xlsx")]
    {
        let (expected, actual) = changed();
        assert_code(
            litchi_xlsx::Error::Package(litchi_opc::OpcError::SourceChanged { expected, actual }),
            Code::IoError,
        );
    }
    #[cfg(feature = "doc")]
    {
        let (expected, observed) = changed();
        assert_code(
            litchi_cfb::OleError::SourceChanged { expected, observed },
            Code::IoError,
        );
    }
    #[cfg(feature = "xls")]
    {
        let (expected, observed) = changed();
        assert_code(
            litchi_xls::SourceBackedError::SourceChanged { expected, observed },
            Code::IoError,
        );
    }
}

// Row: no format recognized -> `unsupported_format`, `detected: None`.
#[test]
fn row_unrecognized_input_is_unsupported_format_without_detection() {
    let error = assert_code(litchi_core::Error::NotOfficeFile, Code::UnsupportedFormat);
    assert_eq!(error.detected(), None);
    let error = assert_code(Error::unsupported_format(None), Code::UnsupportedFormat);
    assert_eq!(error.detected(), None);
    #[cfg(feature = "doc")]
    {
        let error = assert_code(litchi_cfb::OleError::NotOleFile, Code::UnsupportedFormat);
        assert_eq!(error.detected(), None);
    }
}

// Row: wrong family for a view -> `unsupported_format`, `detected: Some(..)`.
#[test]
fn row_wrong_family_is_unsupported_format_with_detection() {
    let error = assert_code(
        litchi_core::Error::UnexpectedFormat {
            detected: FileFormat::Xlsx,
        },
        Code::UnsupportedFormat,
    );
    assert_eq!(error.detected(), Some(FileFormat::Xlsx));
    let error = Error::unsupported_format(Some(FileFormat::Pptx));
    assert_eq!(error.code(), Code::UnsupportedFormat);
    assert_eq!(error.detected(), Some(FileFormat::Pptx));
}

// Row: feature disabled; iWork input in M0a -> `unsupported_format` plus a
// suggestion.
#[test]
fn row_disabled_feature_and_iwork_are_unsupported_format_with_suggestion() {
    assert_suggested(
        litchi_core::Error::FeatureDisabled("xlsb".into()),
        Code::UnsupportedFormat,
        "enable the `xlsb` cargo feature",
    );
    let error = Error::iwork_unsupported(FileFormat::Pages);
    assert_eq!(error.code(), Code::UnsupportedFormat);
    assert_eq!(error.detected(), Some(FileFormat::Pages));
    assert!(
        error
            .suggestion()
            .is_some_and(|suggestion| suggestion.contains("litchi::iwork::Document::open"))
    );
}

// Row: `ParseError`, `InvalidFormat`, `CorruptedFile`, `XmlError`, `ZipError`,
// `InvalidContentType`, malformed leaf payload -> `corrupt_file`.
#[test]
fn row_malformed_input_is_corrupt_file() {
    for leaf in [
        litchi_core::Error::ParseError("bad".into()),
        litchi_core::Error::InvalidFormat("bad".into()),
        litchi_core::Error::CorruptedFile("bad".into()),
        litchi_core::Error::ComponentNotFound("word/document.xml".into()),
        litchi_core::Error::XmlError("bad".into()),
        litchi_core::Error::ZipError("bad".into()),
        litchi_core::Error::InvalidContentType {
            expected: "a".into(),
            got: "b".into(),
        },
    ] {
        assert_code(leaf, Code::CorruptFile);
    }
    #[cfg(feature = "doc")]
    {
        assert_code(
            litchi_cfb::OleError::CorruptedFile("bad".into()),
            Code::CorruptFile,
        );
        assert_code(
            litchi_doc::Error::Ole(litchi_cfb::OleError::StreamNotFound),
            Code::CorruptFile,
        );
    }
    #[cfg(feature = "xls")]
    assert_code(
        litchi_xls::Error::InvalidRecord {
            record_type: 0x0809,
            message: "bad".into(),
        },
        Code::CorruptFile,
    );
    #[cfg(feature = "ppt")]
    assert_code(
        litchi_ppt::Error::Corrupted("bad".into()),
        Code::CorruptFile,
    );
}

// Row: `ResourceLimit`, OPC `ReadLimit`, leaf limit errors -> `limit_exceeded`
// (record 0007, Q1), with the limit detail kept in the typed source.
#[test]
fn row_resource_limits_are_limit_exceeded() {
    let check = |error: Error| {
        assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");
        assert_eq!(error.suggestion(), Some(LIMIT_SUGGESTION), "{error:?}");
        assert!(std::error::Error::source(&error).is_some());
    };
    check(litchi_core::Error::ResourceLimit(resource_limit()).into());
    check(ExecutionError::ResourceLimit(resource_limit()).into());
    #[cfg(feature = "docx")]
    {
        check(
            litchi_opc::OpcError::ReadLimit {
                resource: litchi_opc::ReadResource::ArchiveMembers,
                actual: 11,
                maximum: 10,
            }
            .into(),
        );
        check(
            litchi_docx::Error::RevisionLimit {
                resource: "revisions",
                actual: 11,
                maximum: 10,
            }
            .into(),
        );
    }
    #[cfg(feature = "pptx")]
    check(
        litchi_pptx::Error::Limit {
            resource: "slides",
            limit: 10,
        }
        .into(),
    );
    #[cfg(feature = "xlsx")]
    check(litchi_xlsx::Error::ResourceLimit(resource_limit()).into());
    #[cfg(feature = "xlsb")]
    {
        check(
            litchi_xlsb::package::PackageError::LimitExceeded {
                resource: "cells",
                actual: 11,
                maximum: 10,
            }
            .into(),
        );
        check(
            litchi_xlsb::Error::PayloadLimit {
                length: 11,
                limit: 10,
                offset: 0,
            }
            .into(),
        );
    }
    #[cfg(feature = "doc")]
    {
        check(
            litchi_cfb::OleError::LimitExceeded {
                resource: "input bytes",
                observed: 11,
                maximum: 10,
            }
            .into(),
        );
        let limits = litchi_doc::Limits::default()
            .with_max_package_bytes(1)
            .expect("a one-byte package limit is valid");
        let error =
            litchi_doc::Package::from_reader_with_limits(io::Cursor::new(vec![0_u8; 64]), limits)
                .err()
                .expect("a 64-byte package exceeds a one-byte limit");
        assert!(
            matches!(error, litchi_doc::Error::ResourceLimit(_)),
            "{error:?}"
        );
        check(error.into());
    }
    #[cfg(feature = "xls")]
    {
        check(
            litchi_xls::Error::TooMany {
                collection: "sheets",
                limit: 10,
            }
            .into(),
        );
        check(
            litchi_xls::SourceBackedError::ResourceLimit {
                resource: "records",
                observed: 11,
                maximum: 10,
            }
            .into(),
        );
    }
    #[cfg(feature = "ppt")]
    check(litchi_ppt::Error::ResourceLimit("records".into()).into());
    #[cfg(feature = "rtf")]
    check(
        litchi_rtf::Error::LimitExceeded {
            resource: "groups",
            observed: 11,
            limit: 10,
        }
        .into(),
    );
}

// Row: `Unsupported` (per-format capability) -> `unsupported_capability`.
#[test]
fn row_unsupported_operations_are_unsupported_capability() {
    assert_code(
        litchi_core::Error::Unsupported("sheet selection".into()),
        Code::UnsupportedCapability,
    );
    #[cfg(feature = "xlsx")]
    {
        assert_code(
            litchi_xlsx::Error::Unsupported { feature: "pivots" },
            Code::UnsupportedCapability,
        );
        assert_code(
            litchi_xlsx::Error::UnsupportedSelector,
            Code::UnsupportedCapability,
        );
    }
    #[cfg(feature = "xlsb")]
    assert_code(
        litchi_xlsb::package::PackageError::UnsupportedFeature("decrypt".into()),
        Code::UnsupportedCapability,
    );
    #[cfg(feature = "docx")]
    assert_code(
        litchi_opc::OpcError::PreservationUnavailable {
            reason: "framing".into(),
        },
        Code::UnsupportedCapability,
    );
    #[cfg(feature = "doc")]
    assert_code(
        litchi_doc::Error::UnsupportedEncryption(litchi_doc::EncryptionKind::XorObfuscation),
        Code::UnsupportedCapability,
    );
}

// Row: `Locked` read/edit, `Signed` edit, `ProtectionDenied` ->
// `document_protected` plus a suggestion.
#[test]
fn row_protection_is_document_protected_with_suggestion() {
    let error = Error::locked(Some(FileFormat::Docx));
    assert_eq!(error.code(), Code::DocumentProtected);
    assert_eq!(error.suggestion(), Some(super::LOCKED_SUGGESTION));
    assert_eq!(error.format(), Some(FileFormat::Docx));
    let error = Error::locked(None);
    assert_eq!(error.code(), Code::DocumentProtected);
    assert_eq!(error.format(), None);
    let error = Error::signed(FileFormat::Xlsx);
    assert_eq!(error.code(), Code::DocumentProtected);
    assert_eq!(error.suggestion(), Some(SIGNED_SUGGESTION));
    assert_eq!(error.format(), Some(FileFormat::Xlsx));
    #[cfg(feature = "docx")]
    assert_suggested(
        litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
        Code::DocumentProtected,
        SIGNED_SUGGESTION,
    );
    #[cfg(feature = "xlsx")]
    assert_suggested(
        litchi_xlsx::Error::Signed,
        Code::DocumentProtected,
        SIGNED_SUGGESTION,
    );
    #[cfg(feature = "xlsb")]
    {
        assert_suggested(
            litchi_xlsb::package::PackageError::Signed,
            Code::DocumentProtected,
            SIGNED_SUGGESTION,
        );
        assert_suggested(
            litchi_xlsb::package::PackageError::PasswordProtected,
            Code::DocumentProtected,
            PASSWORD_SUGGESTION,
        );
    }
    #[cfg(feature = "doc")]
    {
        assert_suggested(
            litchi_doc::Error::PasswordRequired,
            Code::DocumentProtected,
            PASSWORD_SUGGESTION,
        );
        let error = assert_code(
            litchi_doc::Error::ProtectionDenied(litchi_doc::EditProtection::Document),
            Code::DocumentProtected,
        );
        assert!(error.suggestion().is_some());
    }
    #[cfg(feature = "xls")]
    {
        assert_suggested(
            litchi_xls::Error::PasswordRequired,
            Code::DocumentProtected,
            PASSWORD_SUGGESTION,
        );
        assert_suggested(
            litchi_xls::Error::PasswordProtected,
            Code::DocumentProtected,
            PASSWORD_SUGGESTION,
        );
        assert_suggested(
            litchi_xls::SourceBackedError::EncryptedUnsupported,
            Code::DocumentProtected,
            PASSWORD_SUGGESTION,
        );
    }
    #[cfg(feature = "ppt")]
    assert_suggested(
        litchi_ppt::Error::PasswordRequired,
        Code::DocumentProtected,
        PASSWORD_SUGGESTION,
    );
}

// Row: `Allocation`, `Committed`, `Other` -> `internal_error`.
#[test]
fn row_allocation_commit_and_other_are_internal_error() {
    assert_code(
        litchi_core::Error::Allocation {
            resource: "buffer",
            source: allocation_failure(),
        },
        Code::InternalError,
    );
    assert_code(
        litchi_core::Error::Committed(io::Error::other("fsync")),
        Code::InternalError,
    );
    assert_code(litchi_core::Error::Other("x".into()), Code::InternalError);
    #[cfg(feature = "docx")]
    {
        assert_code(
            litchi_opc::OpcError::Committed {
                source: io::Error::other("fsync"),
            },
            Code::InternalError,
        );
        assert_code(
            litchi_docx::Error::Allocation {
                resource: "buffer",
                source: allocation_failure(),
            },
            Code::InternalError,
        );
    }
    #[cfg(feature = "pptx")]
    assert_code(
        litchi_pptx::Error::Allocation {
            resource: "buffer",
            source: allocation_failure(),
        },
        Code::InternalError,
    );
    #[cfg(feature = "xlsx")]
    assert_code(
        litchi_xlsx::Error::Allocation {
            resource: "buffer",
            source: allocation_failure(),
        },
        Code::InternalError,
    );
    #[cfg(feature = "doc")]
    assert_code(
        litchi_cfb::OleError::Committed {
            source: io::Error::other("fsync"),
        },
        Code::InternalError,
    );
    #[cfg(feature = "xls")]
    assert_code(litchi_xls::Error::Allocation("cells"), Code::InternalError);
    #[cfg(feature = "ppt")]
    assert_code(
        litchi_ppt::Error::AllocationFailed("records"),
        Code::InternalError,
    );
    #[cfg(feature = "rtf")]
    assert_code(
        litchi_rtf::Error::AllocationFailed {
            resource: "text",
            requested: 1,
        },
        Code::InternalError,
    );
}

// Row: cancellation -> `cancelled` (record 0007, Q1).
#[test]
fn row_cancellation_is_cancelled() {
    assert_code(ExecutionError::Cancelled, Code::Cancelled);
    assert_code(io::Error::other(ExecutionError::Cancelled), Code::Cancelled);
    #[cfg(feature = "docx")]
    {
        assert_code(litchi_opc::OpcError::Cancelled, Code::Cancelled);
        assert_code(
            litchi_opc::OpcError::Execution(ExecutionError::Cancelled),
            Code::Cancelled,
        );
        assert_code(
            litchi_docx::Error::Opc(litchi_opc::OpcError::Cancelled),
            Code::Cancelled,
        );
    }
    #[cfg(feature = "xlsb")]
    assert_code(
        litchi_xlsb::package::PackageError::Execution(ExecutionError::Cancelled),
        Code::Cancelled,
    );
}

// Classifications beyond the table: rejected caller configuration, a wrong
// password, out-of-range indexes, missing names, unsupported versions.
#[test]
fn caller_errors_are_classified_by_kind() {
    #[cfg(feature = "docx")]
    {
        assert_code(
            litchi_opc::OpcError::InvalidReadLimit {
                resource: litchi_opc::ReadResource::ArchiveMembers,
                value: 0,
            },
            Code::InvalidArgument,
        );
        assert_code(
            litchi_docx::Error::OutOfBounds {
                object: "paragraph",
                index: 3,
                len: 2,
            },
            Code::OutOfRange,
        );
    }
    #[cfg(feature = "pptx")]
    {
        assert_code(
            litchi_pptx::Error::SlideIndexOutOfBounds { index: 3, len: 2 },
            Code::OutOfRange,
        );
        assert_code(
            litchi_pptx::Error::SlideNameNotFound("intro".into()),
            Code::NotFound,
        );
    }
    #[cfg(feature = "xlsx")]
    assert_code(
        litchi_xlsx::Error::NotWorksheet {
            sheet: "Chart1".into(),
        },
        Code::UnsupportedType,
    );
    #[cfg(feature = "doc")]
    {
        assert_code(
            litchi_cfb::OleError::InvalidLimit {
                resource: "input bytes",
                value: 0,
                maximum: 1,
            },
            Code::InvalidArgument,
        );
        assert_code(litchi_doc::Error::InvalidPassword, Code::InvalidArgument);
        assert_code(
            litchi_doc::Error::UnsupportedVersion {
                nfib: 0x00c1,
                name: "Word 6",
            },
            Code::UnsupportedFormat,
        );
    }
    #[cfg(feature = "xls")]
    assert_code(
        litchi_xls::Error::UnsupportedBiffVersion(0x0500),
        Code::UnsupportedFormat,
    );
}

// Wildcard arm of the `ExecutionError` classifier: inconsistent policies.
#[test]
fn wildcard_execution_policy_errors_are_invalid_argument() {
    assert_code(
        ExecutionError::WorkersExceedInFlightTasks {
            workers: NonZeroUsize::new(4).expect("non-zero"),
            max_in_flight_tasks: NonZeroUsize::new(1).expect("non-zero"),
        },
        Code::InvalidArgument,
    );
}

// The `litchi_core::Error` classifier names every current variant; its
// wildcard arm only catches variants added later, so it has no reachable
// input to test.

#[cfg(feature = "docx")]
#[test]
fn wildcard_opc_errors_are_corrupt_file() {
    assert_code(
        litchi_opc::OpcError::InvalidPackUri("bad".into()),
        Code::CorruptFile,
    );
}

#[cfg(feature = "docx")]
#[test]
fn wildcard_docx_errors_are_corrupt_file() {
    let error = assert_code(litchi_docx::Error::Xml("bad".into()), Code::CorruptFile);
    assert_eq!(error.format(), Some(FileFormat::Docx));
    assert!(matches!(error.leaf(), Source::Docx(_)));
}

#[cfg(feature = "pptx")]
#[test]
fn wildcard_pptx_errors_are_corrupt_file() {
    let error = assert_code(litchi_pptx::Error::Xml("bad".into()), Code::CorruptFile);
    assert_eq!(error.format(), Some(FileFormat::Pptx));
}

#[cfg(feature = "xlsx")]
#[test]
fn wildcard_xlsx_errors_are_corrupt_file() {
    let error = assert_code(litchi_xlsx::Error::Invalid("bad".into()), Code::CorruptFile);
    assert_eq!(error.format(), Some(FileFormat::Xlsx));
}

#[cfg(feature = "xlsb")]
#[test]
fn wildcard_xlsb_package_errors_are_corrupt_file() {
    let error = assert_code(
        litchi_xlsb::package::PackageError::InvalidRecordType(0x7fff),
        Code::CorruptFile,
    );
    assert_eq!(error.format(), Some(FileFormat::Xlsb));
}

#[cfg(feature = "xlsb")]
#[test]
fn wildcard_xlsb_raw_errors_are_corrupt_file() {
    assert_code(
        litchi_xlsb::Error::InvalidKind { offset: 0 },
        Code::CorruptFile,
    );
}

#[cfg(feature = "doc")]
#[test]
fn wildcard_doc_errors_are_corrupt_file() {
    let error = assert_code(
        litchi_doc::Error::Corrupted("bad".into()),
        Code::CorruptFile,
    );
    assert_eq!(error.format(), Some(FileFormat::Doc));
}

#[cfg(feature = "rtf")]
#[test]
fn wildcard_rtf_errors_are_corrupt_file() {
    let error = assert_code(litchi_rtf::Error::UnexpectedEof, Code::CorruptFile);
    assert_eq!(error.format(), Some(FileFormat::Rtf));
}

#[cfg(feature = "formula")]
#[test]
fn formula_conversion_errors_keep_their_typed_source() {
    use litchi_formula::{FormulaError, LatexError, OmmlError};

    let unsupported = assert_code(
        FormulaError::Omml(OmmlError::UnsupportedFeature("m:groupChr".into())),
        Code::UnsupportedElement,
    );
    assert!(matches!(unsupported.leaf(), Source::Formula(_)));
    assert_code(
        FormulaError::Omml(OmmlError::DepthLimitExceeded(64)),
        Code::LimitExceeded,
    );
    assert_code(
        FormulaError::Omml(OmmlError::ArenaAllocationError("full".into())),
        Code::InternalError,
    );
    // The wildcard arm: every other conversion failure is malformed markup.
    assert_code(
        FormulaError::Omml(OmmlError::MalformedElement("m:f".into())),
        Code::CorruptFile,
    );
    assert_code(
        FormulaError::Latex(LatexError::InvalidNode("bad".into())),
        Code::CorruptFile,
    );
}
