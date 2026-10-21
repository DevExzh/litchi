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

#[test]
fn atomic_destination_conflicts_keep_their_typed_source_and_file_exists_code() {
    let error = Error::from(io::Error::from(io::ErrorKind::AlreadyExists));
    assert_eq!(error.code(), Code::FileExists);
    assert!(matches!(error.0.source.as_ref(), Source::Io(_)));

    #[cfg(feature = "docx")]
    {
        let error = Error::from(litchi_opc::OpcError::IoError(io::Error::from(
            io::ErrorKind::AlreadyExists,
        )));
        assert_eq!(error.code(), Code::FileExists);
        assert!(matches!(error.0.source.as_ref(), Source::Opc(_)));
    }
}

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

#[test]
fn attribution_metadata_defaults_and_round_trips() {
    let plain = Error::new(Code::InvalidPath, "bad path");
    assert_eq!(plain.item(), None);
    assert_eq!(plain.property(), None);

    let error = plain
        .with_item(17)
        .with_property(String::from("font.bolld"));
    assert_eq!(error.item(), Some(17));
    assert_eq!(error.property(), Some("font.bolld"));
    assert_eq!(
        error.to_string(),
        "invalid_path: bad path (property: font.bolld) (item: 17)"
    );
}

#[test]
fn attribution_metadata_preserves_classification_source_and_clone() {
    let error = Error::from(litchi_core::Error::ResourceLimit(resource_limit()))
        .with_item(usize::MAX)
        .with_property("worksheet.cells");
    let clone = error.clone();

    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(clone.code(), Code::LimitExceeded);
    assert_eq!(clone.item(), Some(usize::MAX));
    assert_eq!(clone.property(), Some("worksheet.cells"));
    assert!(matches!(error.leaf(), Source::Core(_)));
    assert!(std::ptr::eq(error.leaf(), clone.leaf()));
}

#[test]
fn debug_redacts_authored_messages_and_attribution_payloads() {
    let secret = "private-property-name";
    let error = Error::new(
        Code::UnsupportedProperty,
        format!("unknown property {secret}"),
    )
    .with_item(9_876_543)
    .with_property(secret.to_owned())
    .with_suggestion(format!("try {secret}"));
    let debug = format!("{error:?}");

    assert!(
        !debug.contains(secret),
        "debug leaked authored input: {debug}"
    );
    assert!(
        !debug.contains("9_876_543"),
        "debug leaked the item attribution: {debug}"
    );
    assert!(debug.contains("UnsupportedProperty"));
    assert_eq!(error.property(), Some(secret));
    assert_eq!(error.item(), Some(9_876_543));
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
            .is_some_and(|suggestion| suggestion.contains("`pages`, `keynote` and `numbers`"))
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

// Row: `ResourceLimit`, OPC `Limit`, leaf limit errors -> `limit_exceeded`
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
            litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::Members,
                observed: 11,
                maximum: 10,
            }
            .into(),
        );
        check(
            litchi_docx::Error::RevisionLimit {
                kind: litchi_docx::revision::LimitKind::Records,
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
        let limits = litchi_doc::SourceLimits::builder()
            .max_input_bytes(1)
            .unwrap()
            .build()
            .expect("a one-byte package limit is valid");
        let error = litchi_doc::Package::from_reader_with_options(
            io::Cursor::new(vec![0_u8; 64]),
            litchi_doc::OpenOptions::default().source_limits(limits),
        )
        .expect_err("a 64-byte package exceeds a one-byte limit");
        assert!(
            matches!(error, litchi_doc::Error::Limit { .. }),
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
    check(
        litchi_ppt::Error::Limit {
            kind: litchi_ppt::SourceLimitKind::Events,
            observed: 11,
            maximum: 10,
        }
        .into(),
    );
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
    // An `\ansicpg` without an exact codec (record 0347).
    #[cfg(feature = "rtf")]
    assert_code(
        litchi_rtf::Error::UnsupportedCodePage(10001),
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
            litchi_opc::OpcError::InvalidLimit {
                kind: litchi_opc::SourceLimitKind::Members,
                value: 0,
                maximum: usize::MAX as u64,
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
        assert_code(
            litchi_pptx::Error::SlideLayoutNameNotFound("Blank".into()),
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

// The source-backed legacy editors and the compound-file signature editor
// (record 0013). Their classifiers name every current variant of
// `litchi_doc::source_backed::Error`, `litchi_cfb::OverlayError` and
// `litchi_sign::Error`, so those wildcard arms only catch variants added
// later; the refusal wildcards below are reachable.

#[cfg(feature = "doc")]
#[test]
fn doc_body_text_errors_are_classified() {
    use litchi_cfb::OverlayError;
    use litchi_doc::source_backed::{Error as BodyError, Refusal};

    let error = assert_code(BodyError::Refused(Refusal::Signed), Code::DocumentProtected);
    assert_eq!(error.suggestion(), Some(SIGNED_SUGGESTION));
    assert_eq!(error.format(), Some(FileFormat::Doc));
    assert!(matches!(error.leaf(), Source::DocBodyText(_)));
    assert_suggested(
        BodyError::Refused(Refusal::Encrypted),
        Code::DocumentProtected,
        PASSWORD_SUGGESTION,
    );
    assert_code(
        BodyError::Refused(Refusal::Protected),
        Code::DocumentProtected,
    );
    assert_code(
        BodyError::Refused(Refusal::ParagraphNotFound),
        Code::OutOfRange,
    );
    // Wildcard refusal arm.
    assert_code(
        BodyError::Refused(Refusal::LengthChange),
        Code::UnsupportedCapability,
    );
    assert_code(BodyError::InvalidData("bad".into()), Code::CorruptFile);
    assert_code(BodyError::Limit("big".into()), Code::LimitExceeded);
    assert_code(
        BodyError::Allocation {
            resource: "test",
            source: allocation_failure(),
        },
        Code::InternalError,
    );
    assert_code(BodyError::Io(io::Error::other("sink")), Code::IoError);
    assert_code(
        BodyError::Overlay(OverlayError::Unavailable {
            reason: "test".into(),
        }),
        Code::UnsupportedCapability,
    );
    let (expected, observed) = changed();
    assert_code(
        BodyError::Overlay(OverlayError::SourceChanged { expected, observed }),
        Code::IoError,
    );
}

#[cfg(feature = "doc")]
#[test]
fn doc_edit_errors_keep_the_edit_type_and_map_native_categories() {
    use litchi_doc::body_text::{Error as EditError, Refusal};
    use litchi_doc::package::Error as PackageError;

    let protected = assert_code(
        EditError::Refused(Refusal::Protected),
        Code::DocumentProtected,
    );
    assert!(matches!(protected.leaf(), Source::DocEdit(_)));
    assert!(protected.suggestion().is_some());

    assert_code(
        EditError::Refused(Refusal::OperationLimit {
            observed: 2,
            limit: 1,
        }),
        Code::LimitExceeded,
    );
    assert_code(
        EditError::Refused(Refusal::ResourceCollision { storage_id: 7 }),
        Code::DuplicateName,
    );
    assert_code(
        EditError::Refused(Refusal::ResourceNotFound { storage_id: 7 }),
        Code::NotFound,
    );
    assert_code(EditError::Conflict, Code::InvalidInput);
    assert_code(
        EditError::Refused(Refusal::StructuralContent),
        Code::UnsupportedCapability,
    );

    let invalid = assert_code(
        EditError::Invalid(PackageError::Corrupted(
            "authored body payload must remain private".into(),
        )),
        Code::CorruptFile,
    );
    assert!(matches!(invalid.leaf(), Source::DocEdit(_)));
    assert!(std::error::Error::source(&invalid).is_some());
}

#[cfg(feature = "doc")]
#[test]
fn doc_edit_diagnostics_are_redacted_but_retain_the_typed_chain() {
    use litchi_doc::body_text::{Error as EditError, Refusal};
    use litchi_doc::package::Error as PackageError;

    let secret = "private authored DOC text";
    let error = Error::from(EditError::Invalid(PackageError::Corrupted(
        secret.to_owned(),
    )));
    for diagnostic in [
        error.message().to_owned(),
        error.to_string(),
        format!("{error:?}"),
        format!("{:?}", error.leaf()),
    ] {
        assert!(!diagnostic.contains(secret), "diagnostic leaked {secret:?}");
    }
    assert!(matches!(error.leaf(), Source::DocEdit(_)));
    assert!(std::error::Error::source(&error).is_some());

    let protected = Error::from(EditError::Refused(Refusal::Protected));
    assert!(matches!(protected.leaf(), Source::DocEdit(_)));
}

#[cfg(feature = "doc")]
#[test]
fn doc_write_errors_keep_the_write_type_and_map_native_categories() {
    use litchi_doc::package::Error as PackageError;
    use litchi_doc::writer::WriteError;

    let unsupported = assert_code(WriteError::EncryptedMetadata, Code::UnsupportedCapability);
    assert!(matches!(unsupported.leaf(), Source::DocWrite(_)));
    assert!(std::error::Error::source(&unsupported).is_some());
    assert_code(
        WriteError::Io(io::Error::from(io::ErrorKind::NotFound)),
        Code::FileNotFound,
    );
    assert_code(
        WriteError::InvalidData("invalid authored text".into()),
        Code::CorruptFile,
    );
    assert_code(
        WriteError::Ole(litchi_cfb::OleError::InvalidFormat("bad OLE".into())),
        Code::CorruptFile,
    );
    assert_code(
        WriteError::Package(PackageError::Corrupted("bad package".into())),
        Code::CorruptFile,
    );
}

#[cfg(feature = "doc")]
#[test]
fn doc_write_diagnostics_are_redacted_but_retain_the_typed_chain() {
    use litchi_doc::writer::WriteError;

    let secret = "private authored DOC property";
    let error = Error::from(WriteError::InvalidData(secret.to_owned()));
    for diagnostic in [
        error.message().to_owned(),
        error.to_string(),
        format!("{error:?}"),
        format!("{:?}", error.leaf()),
    ] {
        assert!(!diagnostic.contains(secret), "diagnostic leaked {secret:?}");
    }
    assert!(matches!(error.leaf(), Source::DocWrite(_)));
    assert!(std::error::Error::source(&error).is_some());
}

#[cfg(feature = "ppt")]
#[test]
fn ppt_text_edit_errors_are_classified() {
    use litchi_cfb::OverlayError;
    use litchi_ppt::text_edit::{Error as TextError, Refusal};

    assert_code(
        TextError::Refused(Refusal::SlideNotFound {
            position: litchi_core::Position::new(3),
        }),
        Code::NotFound,
    );
    assert_code(TextError::Refused(Refusal::ShapeNotFound), Code::NotFound);
    // Classified by kind: the protected-container refusal is unsupported
    // content, a shape without a textbox an unsupported target.
    assert_code(
        TextError::Refused(Refusal::UnsupportedSource),
        Code::UnsupportedCapability,
    );
    assert_code(
        TextError::Refused(Refusal::NoTextbox),
        Code::UnsupportedType,
    );
    assert_code(
        TextError::Package(litchi_ppt::Error::PasswordRequired),
        Code::DocumentProtected,
    );
    assert_code(TextError::PatchConflict, Code::InvalidInput);
    assert_code(TextError::Signed, Code::DocumentProtected);
    let error = assert_code(
        TextError::Source(OverlayError::Allocation {
            resource: "test",
            source: allocation_failure(),
        }),
        Code::InternalError,
    );
    assert_eq!(error.format(), Some(FileFormat::Ppt));
    assert!(matches!(error.leaf(), Source::PptTextEdit(_)));
}

#[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
#[test]
fn compound_file_signature_errors_are_classified() {
    let error = assert_code(litchi_sign::Error::Xml("bad".into()), Code::CorruptFile);
    assert_eq!(error.format(), None);
    assert!(matches!(error.leaf(), Source::Sign(_)));
    assert_code(
        litchi_sign::Error::Container("bad".into()),
        Code::CorruptFile,
    );
    assert_code(litchi_sign::Error::Limit("big".into()), Code::LimitExceeded);
    assert_code(litchi_sign::Error::Encrypted, Code::UnsupportedCapability);
    assert_code(litchi_sign::Error::Legacy, Code::UnsupportedCapability);
    assert_code(litchi_sign::Error::Key("bad".into()), Code::InvalidArgument);
    assert_code(litchi_sign::Error::Sign("bad".into()), Code::InternalError);
    assert_code(
        litchi_sign::Error::Cfb(litchi_cfb::OleError::NotOleFile),
        Code::UnsupportedFormat,
    );
}

#[cfg(feature = "docx")]
#[test]
fn docx_owning_policy_refusal_is_an_unsupported_capability() {
    let error = assert_code(
        litchi_docx::Error::UnsupportedOpenPolicy {
            policy: litchi_docx::OpenPolicy::ExecutionContext,
        },
        Code::UnsupportedCapability,
    );
    assert!(matches!(error.leaf(), Source::Docx(_)));
}

#[cfg(all(feature = "docx", feature = "encryption"))]
#[test]
fn docx_encryption_policy_refusal_retains_typed_source() {
    let error = assert_code(
        litchi_docx::Error::EncryptionPolicy {
            operation: "to_bytes",
            source: litchi_ooxml_common::package_encryption::PolicyError::NoRetainedMode,
        },
        Code::UnsupportedCapability,
    );
    assert!(matches!(
        error.leaf(),
        Source::Docx(litchi_docx::Error::EncryptionPolicy { .. })
    ));
}

#[cfg(feature = "xlsb")]
#[test]
fn xlsb_diagnostics_hide_authored_payloads_but_retain_typed_causes() {
    let secret = "private workbook formula and sheet name";
    let error = Error::from(litchi_xlsb::package::PackageError::InvalidFormat(
        secret.into(),
    ));
    for diagnostic in [
        error.message().to_owned(),
        error.to_string(),
        format!("{error:?}"),
        format!("{:?}", error.leaf()),
    ] {
        assert!(!diagnostic.contains(secret));
    }
    assert!(matches!(error.leaf(), Source::Xlsb(_)));
    assert!(std::error::Error::source(&error).is_some());
}

#[cfg(feature = "docx")]
#[test]
fn docx_schema_property_refusal_is_an_invalid_value() {
    let error = Error::from(litchi_docx::document::Error::Refused {
        position: 0,
        reason: litchi_docx::document::Refusal::InvalidPropertyValue,
    });
    assert_eq!(error.code(), Code::InvalidValue);
    assert!(matches!(error.leaf(), Source::DocxTransaction(_)));
}

#[cfg(all(feature = "xlsx", feature = "encryption"))]
#[test]
fn xlsx_output_policy_retains_the_typed_cause() {
    use litchi_ooxml_common::package_encryption::PolicyError;
    let error = assert_code(
        litchi_xlsx::Error::EncryptionPolicy {
            operation: "write_to",
            source: PolicyError::OrdinaryOutputWouldRemoveEncryption {
                mode: litchi_xlsx::encryption::Mode::Agile,
            },
        },
        Code::DocumentProtected,
    );
    assert!(matches!(
        error.leaf(),
        Source::Xlsx(litchi_xlsx::Error::EncryptionPolicy { .. })
    ));
    assert!(error.suggestion().is_some());
    assert_code(
        litchi_xlsx::Error::EncryptionPolicy {
            operation: "to_reencrypted",
            source: PolicyError::NoRetainedMode,
        },
        Code::UnsupportedCapability,
    );
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_edit_guards_preserve_protection_and_capability_categories() {
    use litchi_xlsx::{ColumnEditBlock, EditBlock, RowEditBlock};
    for (native, code) in [
        (
            litchi_xlsx::Error::RowEditBlocked {
                sheet: "Sheet1".into(),
                row: litchi_xlsx::RowIndex::new(0).unwrap(),
                reason: RowEditBlock::ProtectedSheet,
            },
            Code::DocumentProtected,
        ),
        (
            litchi_xlsx::Error::ColumnEditBlocked {
                sheet: "Sheet1".into(),
                column: litchi_xlsx::ColumnIndex::new(0).unwrap(),
                reason: ColumnEditBlock::ProtectedSheet,
            },
            Code::DocumentProtected,
        ),
        (
            litchi_xlsx::Error::ColumnEditBlocked {
                sheet: "Sheet1".into(),
                column: litchi_xlsx::ColumnIndex::new(0).unwrap(),
                reason: ColumnEditBlock::MarkupCompatibility,
            },
            Code::UnsupportedCapability,
        ),
        (
            litchi_xlsx::Error::EditBlocked {
                sheet: "Sheet1".into(),
                address: litchi_xlsx::Address::new(
                    litchi_xlsx::RowIndex::new(0).unwrap(),
                    litchi_xlsx::ColumnIndex::new(0).unwrap(),
                ),
                reason: EditBlock::ProtectedSheet,
            },
            Code::DocumentProtected,
        ),
        (
            litchi_xlsx::Error::EditBlocked {
                sheet: "Sheet1".into(),
                address: litchi_xlsx::Address::new(
                    litchi_xlsx::RowIndex::new(0).unwrap(),
                    litchi_xlsx::ColumnIndex::new(0).unwrap(),
                ),
                reason: EditBlock::CoveredMerge,
            },
            Code::UnsupportedCapability,
        ),
    ] {
        let error = assert_code(native, code);
        assert!(matches!(error.leaf(), Source::Xlsx(_)));
    }
}

#[cfg(all(feature = "pptx", feature = "encryption"))]
#[test]
fn pptx_crypto_errors_keep_their_classification() {
    let error = Error::from(litchi_pptx::Error::Encryption(
        litchi_pptx::encryption::Error::Limit {
            resource: "plaintext",
            actual: 2,
            maximum: 1,
        },
    ));
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(error.format(), Some(FileFormat::Pptx));
    let error = Error::from(litchi_pptx::Error::Encryption(
        litchi_pptx::encryption::Error::PasswordRequired,
    ));
    assert_eq!(error.code(), Code::InvalidArgument);
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_slide_removal_refusals_preserve_semantic_classification() {
    use litchi_pptx::SlideRemovalRefusal as Refusal;
    for (kind, code) in [
        (Refusal::SignedPackage, Code::DocumentProtected),
        (Refusal::ProtectedPresentation, Code::DocumentProtected),
        (Refusal::FinalSlide, Code::InvalidArgument),
        (Refusal::MacroEnabledPackage, Code::UnsupportedCapability),
        (Refusal::MarkupCompatibility, Code::UnsupportedCapability),
        (Refusal::UnknownSemanticSurface, Code::UnsupportedCapability),
        (Refusal::SharedOwner, Code::UnsupportedCapability),
        (
            Refusal::UnsupportedRelationship,
            Code::UnsupportedCapability,
        ),
        (Refusal::AmbiguousTopology, Code::UnsupportedCapability),
        (Refusal::UnknownPhysicalMember, Code::UnsupportedCapability),
    ] {
        assert_code(
            litchi_pptx::Error::SlideRemovalPlan {
                kind,
                detail: "removal guard".into(),
            },
            code,
        );
    }
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_slide_order_refusals_distinguish_protection_capability_and_corruption() {
    use litchi_pptx::SlideOrderRefusal as Refusal;
    for (kind, code) in [
        (Refusal::ModifyVerifier, Code::DocumentProtected),
        (Refusal::AmbiguousBindings, Code::CorruptFile),
        (Refusal::MarkupCompatibility, Code::UnsupportedCapability),
        (Refusal::Sections, Code::UnsupportedCapability),
        (Refusal::ViewProperties, Code::UnsupportedCapability),
        (Refusal::PresentationProperties, Code::UnsupportedCapability),
        (Refusal::PositionalOwner, Code::UnsupportedCapability),
    ] {
        assert_code(
            litchi_pptx::Error::SlideOrderPlan {
                kind,
                detail: "order guard".into(),
            },
            code,
        );
    }
}

#[cfg(feature = "doc")]
#[test]
fn doc_primary_transactions_preserve_typed_causes_and_safe_messages() {
    use litchi_doc::package::edit::{Error as Native, LimitKind};
    for (native, expected) in [
        (
            Native::Limit {
                kind: LimitKind::Operations,
                observed: 2,
                maximum: 1,
            },
            Code::LimitExceeded,
        ),
        (
            Native::Body(litchi_doc::body_text::Error::Conflict),
            Code::InvalidInput,
        ),
        (
            Native::Body(litchi_doc::body_text::Error::Refused(
                litchi_doc::body_text::Refusal::Protected,
            )),
            Code::DocumentProtected,
        ),
        (
            Native::Package(litchi_doc::Error::PasswordRequired),
            Code::DocumentProtected,
        ),
    ] {
        let error = assert_code(native, expected);
        assert!(matches!(error.0.source.as_ref(), Source::DocTransaction(_)));
        assert_eq!(error.message(), expected.summary());
    }
}

#[cfg(feature = "doc")]
#[test]
fn doc_source_limit_configuration_preserves_dimension_and_values() {
    let native = litchi_doc::SourceLimits::builder()
        .max_input_bytes(litchi_doc::SourceLimits::MAX_INPUT_BYTES + 1)
        .unwrap_err();
    assert!(matches!(
        native,
        litchi_doc::Error::InvalidLimit {
            kind: litchi_doc::SourceLimitKind::InputBytes,
            ..
        }
    ));
    let error = assert_code(native, Code::InvalidArgument);
    assert!(matches!(
        error.leaf(),
        Source::Doc(litchi_doc::Error::InvalidLimit { .. })
    ));
}

#[cfg(feature = "xls")]
#[test]
fn xls_source_limits_preserve_error_codes() {
    assert_code(
        litchi_xls::Error::InvalidLimit {
            kind: litchi_xls::LimitKind::InputBytes,
            value: 0,
            maximum: litchi_xls::SourceLimits::MAX_INPUT_BYTES as u64,
        },
        Code::InvalidArgument,
    );
    assert_code(
        litchi_xls::Error::Limit {
            kind: litchi_xls::LimitKind::EntryBytes,
            observed: 101,
            maximum: 100,
        },
        Code::LimitExceeded,
    );
}

#[cfg(feature = "xls")]
#[test]
fn xls_package_errors_preserve_classification() {
    assert_code(
        litchi_xls::package::Error::EncryptionPolicy,
        Code::DocumentProtected,
    );
    assert_code(
        litchi_xls::package::Error::Reader(litchi_xls::Error::Limit {
            kind: litchi_xls::LimitKind::EntryBytes,
            observed: 101,
            maximum: 100,
        }),
        Code::LimitExceeded,
    );
}

#[cfg(feature = "xls")]
#[test]
fn xls_staged_changes_classify_as_a_limit() {
    assert_code(
        litchi_xls::Error::Limit {
            kind: litchi_xls::LimitKind::Operations,
            observed: 4097,
            maximum: 4096,
        },
        Code::LimitExceeded,
    );
}

#[cfg(feature = "xls")]
#[test]
fn xls_primary_edit_errors_keep_typed_sources_and_codes() {
    use litchi_xls::package::edit::{Error as EditError, LimitKind};
    assert_code(
        EditError::Limit {
            kind: LimitKind::Operations,
            observed: 4097,
            maximum: 4096,
        },
        Code::LimitExceeded,
    );
    assert_code(EditError::PatchConflict, Code::InvalidInput);
    let error: Error = EditError::Package(litchi_xls::package::Error::EncryptionPolicy).into();
    assert_eq!(error.code(), Code::DocumentProtected);
    assert!(matches!(
        error.leaf(),
        Source::XlsEdit(EditError::Package(
            litchi_xls::package::Error::EncryptionPolicy
        ))
    ));
}

#[cfg(feature = "ppt")]
#[test]
fn ppt_source_limit_configuration_preserves_dimension_and_values() {
    let native = litchi_ppt::SourceLimits::builder()
        .max_input_bytes(litchi_ppt::SourceLimits::MAX_INPUT_BYTES + 1)
        .unwrap_err();
    let error = assert_code(native, Code::InvalidArgument);
    assert!(matches!(
        error.leaf(),
        Source::Ppt(litchi_ppt::Error::InvalidLimit {
            kind: litchi_ppt::SourceLimitKind::InputBytes,
            value,
            maximum,
        }) if *value == litchi_ppt::SourceLimits::MAX_INPUT_BYTES + 1
            && *maximum == litchi_ppt::SourceLimits::MAX_INPUT_BYTES
    ));
    assert_code(
        litchi_ppt::Error::ResourceLimit("domain-specific bound".into()),
        Code::LimitExceeded,
    );
}

#[cfg(feature = "ppt")]
#[test]
fn ppt_authoring_errors_preserve_their_typed_cause() {
    let error = assert_code(
        litchi_ppt::Error::Authoring(litchi_ppt::WriteError::InvalidData(
            "authored-secret".into(),
        )),
        Code::InvalidArgument,
    );
    assert!(matches!(
        error.leaf(),
        Source::Ppt(litchi_ppt::Error::Authoring(
            litchi_ppt::WriteError::InvalidData(_)
        ))
    ));
    assert!(!format!("{error:?} {error}").contains("authored-secret"));
    #[cfg(feature = "encryption")]
    assert_code(
        litchi_ppt::Error::Authoring(litchi_ppt::WriteError::EncryptionEntropyRequired),
        Code::InvalidArgument,
    );
}

#[cfg(feature = "ppt")]
#[test]
fn ppt_primary_edit_errors_preserve_policy_codes_and_typed_sources() {
    use litchi_ppt::package::edit::{Error as EditError, LimitKind};
    let error = assert_code(EditError::Signed, Code::DocumentProtected);
    assert!(matches!(error.leaf(), Source::PptEdit(EditError::Signed)));
    assert_code(EditError::PatchConflict, Code::InvalidInput);
    assert_code(
        EditError::from(litchi_ppt::slide_order::Error::Refused(
            litchi_ppt::slide_order::Refusal::TransferTargetMismatch,
        )),
        Code::InvalidInput,
    );
    assert_code(
        EditError::from(litchi_ppt::slide_order::Error::Refused(
            litchi_ppt::slide_order::Refusal::DuplicateShapeTextTarget {
                target: litchi_ppt::text_edit::Target::new(
                    litchi_core::Position::new(0),
                    litchi_core::Position::new(0),
                ),
            },
        )),
        Code::InvalidValue,
    );

    assert_code(
        EditError::InvalidLimit {
            kind: LimitKind::Operations,
            value: 0,
            maximum: 4096,
        },
        Code::InvalidArgument,
    );
    assert_code(
        EditError::Limit {
            kind: LimitKind::Operations,
            observed: 2,
            maximum: 1,
        },
        Code::LimitExceeded,
    );
}

#[test]
fn edit_refusals_classify_by_kind_alone() {
    use litchi_core::edit::{Refusal, RefusalKind};

    use super::Classify as _;

    let expected = [
        (
            RefusalKind::UnsupportedOperation,
            Code::UnsupportedCapability,
        ),
        (RefusalKind::UnsupportedContent, Code::UnsupportedCapability),
        (RefusalKind::UnsupportedTarget, Code::UnsupportedType),
        (RefusalKind::UnsupportedProperty, Code::UnsupportedProperty),
        (RefusalKind::InvalidValue, Code::InvalidValue),
        (RefusalKind::NotFound, Code::NotFound),
        (RefusalKind::Protected, Code::DocumentProtected),
    ];
    assert_eq!(expected.len(), RefusalKind::ALL.len());
    for (kind, code) in expected {
        assert_eq!(kind.classify().code, code, "{kind}");
        let refusal = Refusal::new(kind, "operation", "reason");
        assert_eq!(refusal.classify().code, code, "{kind}");
    }
}

// Record 0358 (W6): ODF refusals classify by kind alone, a patch made for
// another base is the caller's input, and a changed write of a signed
// package is protected.
#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
#[test]
fn odf_refusals_classify_by_kind_and_patch_conflicts_are_invalid_input() {
    use litchi_core::edit::RefusalKind;
    use litchi_odf_common::Error as OdfError;

    use super::Classify as _;

    for kind in RefusalKind::ALL.iter().copied() {
        let error = OdfError::refused(kind, "operation", "reason");
        assert_eq!(error.classify().code, kind.classify().code, "{kind}");
    }
    assert_eq!(
        OdfError::patch_conflict("operation", "reason")
            .classify()
            .code,
        Code::InvalidInput
    );
    assert_eq!(OdfError::Signed.classify().code, Code::DocumentProtected);
}

// Record 0351 (W1): DOCX refusals classify by kind alone, a patch made for
// another base is the caller's input, and a broken invariant is internal.
#[cfg(feature = "docx")]
#[test]
fn docx_refusals_classify_by_kind_and_patch_conflicts_are_invalid_input() {
    use litchi_core::edit::{Refusal, RefusalKind};

    for (kind, code) in [
        (
            RefusalKind::UnsupportedOperation,
            Code::UnsupportedCapability,
        ),
        (RefusalKind::UnsupportedContent, Code::UnsupportedCapability),
        (RefusalKind::UnsupportedTarget, Code::UnsupportedType),
        (RefusalKind::NotFound, Code::NotFound),
        (RefusalKind::Protected, Code::DocumentProtected),
    ] {
        let error = assert_code(
            litchi_docx::Error::Refused(Refusal::new(kind, "edit", "reason")),
            code,
        );
        assert!(matches!(
            error.leaf(),
            Source::Docx(litchi_docx::Error::Refused(_))
        ));
    }
    assert_code(
        litchi_docx::Error::PatchConflict {
            operation: "apply",
            reason: "another base",
        },
        Code::InvalidInput,
    );
    assert_code(
        litchi_docx::Error::Internal {
            operation: "edit",
            reason: "invariant",
        },
        Code::InternalError,
    );
    assert_code(
        litchi_docx::document::Error::PatchConflict,
        Code::InvalidInput,
    );
    assert_code(
        litchi_docx::metadata::Error::PatchConflict,
        Code::InvalidInput,
    );
    // Record 0355a: document refusal reasons classify by their kind.
    for (reason, code) in [
        (litchi_docx::document::Refusal::RunNotFound, Code::NotFound),
        (
            litchi_docx::document::Refusal::SectionBreak,
            Code::UnsupportedType,
        ),
        (
            litchi_docx::document::Refusal::ComplexRun,
            Code::UnsupportedCapability,
        ),
        (
            litchi_docx::document::Refusal::UnmodelledProperty,
            Code::UnsupportedProperty,
        ),
        (
            litchi_docx::document::Refusal::StructuralText,
            Code::UnsupportedCapability,
        ),
    ] {
        assert_code(
            litchi_docx::document::Error::Refused {
                position: 0,
                reason,
            },
            code,
        );
    }
    assert_code(
        litchi_docx::document::Error::Refused {
            position: 0,
            reason: litchi_docx::document::Refusal::InvalidPropertyValue,
        },
        Code::InvalidValue,
    );
}

#[cfg(feature = "xls")]
#[test]
fn xls_ole_budget_errors_preserve_typed_dimensions_and_codes() {
    use litchi_xls::ole_object::{EmbeddedPayload, LimitKind, Limits, Snapshot};
    let native = EmbeddedPayload::new(
        b"private payload bytes".to_vec(),
        Limits {
            max_object_size: 1,
            ..Limits::default()
        },
    )
    .expect_err("budget admission precedes CFB parsing");
    assert!(matches!(
        native,
        litchi_xls::Error::Limit {
            kind: litchi_xls::LimitKind::OleObject(LimitKind::ObjectBytes),
            observed: 21,
            maximum: 1,
        }
    ));
    assert_code(native, Code::LimitExceeded);

    let native = Snapshot::open(
        Vec::new(),
        Limits {
            max_stream_size: 0,
            ..Limits::default()
        },
    )
    .expect_err("invalid read policy precedes CFB parsing");
    assert!(matches!(
        native,
        litchi_xls::Error::InvalidLimit {
            kind: litchi_xls::LimitKind::OleObject(LimitKind::StreamBytes),
            value: 0,
            ..
        }
    ));
    assert_code(native, Code::InvalidArgument);
}

#[cfg(any(feature = "xls", feature = "ppt"))]
#[test]
fn graph_budget_errors_keep_dimensions_and_host_classification() {
    use litchi_ograph::{Error as GraphError, LimitKind, Limits, OpenOptions, Package};
    let exceeded = || {
        let error = Package::open_with_options(
            vec![0x55; 2],
            OpenOptions::new().limits(Limits {
                max_package_bytes: 1,
                ..Limits::default()
            }),
        )
        .expect_err("byte budget is checked before parsing CFB");
        assert!(matches!(
            error,
            GraphError::Limit {
                kind: LimitKind::PackageBytes,
                observed: 2,
                maximum: 1,
            }
        ));
        error
    };
    let invalid = || {
        let error = Package::open_with_options(
            Vec::new(),
            OpenOptions::new().limits(Limits {
                max_groups: 11,
                ..Limits::default()
            }),
        )
        .expect_err("invalid bounds are checked before parsing CFB");
        assert!(matches!(
            error,
            GraphError::InvalidLimit {
                kind: LimitKind::Groups,
                value: 11,
                maximum: 10,
            }
        ));
        error
    };
    #[cfg(feature = "xls")]
    {
        assert_code(litchi_xls::Error::from(exceeded()), Code::LimitExceeded);
        assert_code(litchi_xls::Error::from(invalid()), Code::InvalidArgument);
    }
    #[cfg(feature = "ppt")]
    {
        assert_code(litchi_ppt::Error::from(exceeded()), Code::LimitExceeded);
        assert_code(litchi_ppt::Error::from(invalid()), Code::InvalidArgument);
        assert_code(
            litchi_ppt::Error::Authoring(litchi_ppt::writer::WriteError::from(exceeded())),
            Code::LimitExceeded,
        );
        assert_code(
            litchi_ppt::Error::Authoring(litchi_ppt::writer::WriteError::from(invalid())),
            Code::InvalidArgument,
        );
    }
}

#[cfg(any(feature = "xls", feature = "ppt"))]
#[test]
fn graph_typed_causes_are_classified_without_parsing_messages() {
    use super::Classify as _;
    use litchi_ograph::Error as GraphError;
    let cases = [
        (
            GraphError::PatchConflict {
                operation: "edit chart",
                reason: "source differs",
            },
            Code::InvalidInput,
        ),
        (
            GraphError::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "edit chart",
                "target is unsupported",
            )),
            Code::UnsupportedType,
        ),
        (
            GraphError::Allocation { resource: "chart" },
            Code::InternalError,
        ),
        (
            GraphError::Biff(litchi_biff::Error::Allocation {
                resource: litchi_biff::Resource::RecordFrame,
            }),
            Code::InternalError,
        ),
        (
            GraphError::Biff(litchi_biff::Error::TruncatedHeader {
                offset: 0,
                available: 2,
            }),
            Code::CorruptFile,
        ),
        (
            GraphError::from(litchi_biff::Error::LimitExceeded {
                resource: litchi_biff::Resource::RecordCount,
                observed: 3,
                maximum: 2,
            }),
            Code::LimitExceeded,
        ),
    ];
    for (error, expected) in cases {
        assert_eq!(error.classify().code, expected);
    }
}
