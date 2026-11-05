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
#[cfg(feature = "xlsx")]
use super::classify_xlsx_join_source;
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

/// A real XLSX join wrapper still takes the explicit join-classification path;
/// non-normalization join failures keep their historical corrupt-file code.
#[test]
#[cfg(feature = "xlsx")]
fn xlsx_join_error_preserves_other_join_fallback_through_ops_conversion() {
    let mut left = litchi_xlsx::Package::new()
        .expect("left workbook")
        .edit()
        .expect("left edit");
    let right = litchi_xlsx::Package::new()
        .expect("right workbook")
        .edit()
        .expect("right edit");
    let join = left
        .join(right)
        .expect_err("edits from different snapshots must not join");

    let leaf = litchi_xlsx::Error::from(join);
    let error: Error = leaf.into();
    assert_eq!(error.code(), Code::CorruptFile);
    assert_eq!(error.format(), Some(FileFormat::Xlsx));
    assert!(matches!(error.leaf(), Source::Xlsx(_)));
    let source = std::error::Error::source(&error).expect("typed XLSX join source");
    assert!(
        source
            .downcast_ref::<litchi_xlsx::Error>()
            .is_some_and(|error| matches!(error, litchi_xlsx::Error::Join(_)))
    );
}

/// The two operation-boundary leaves remain typed when a join retains them;
/// the public cross-snapshot test above covers the no-source join fallback.
#[test]
#[cfg(feature = "xlsx")]
fn xlsx_join_source_classifier_preserves_operation_boundary_codes() {
    let (expected, observed) = changed();
    let source_changed = litchi_xlsx::Error::SourceChanged { expected, observed };
    assert_eq!(
        classify_xlsx_join_source(Some(&source_changed)).code,
        Code::IoError
    );

    let internal = litchi_xlsx::Error::Internal("generated staging invariant");
    assert_eq!(
        classify_xlsx_join_source(Some(&internal)).code,
        Code::InternalError
    );
}

#[test]
#[cfg(feature = "xlsx")]
fn xlsx_join_source_classifier_delegates_retained_leaf_codes() {
    use super::Classify as _;
    use litchi_core::edit::{Refusal, RefusalKind};

    for kind in RefusalKind::ALL.iter().copied() {
        let source = litchi_xlsx::Error::Refused(Refusal::new(kind, "operation", "reason"));
        assert_eq!(
            classify_xlsx_join_source(Some(&source)).code,
            kind.classify().code,
            "{kind}"
        );
    }

    let conflict = litchi_xlsx::Error::PatchConflict {
        part: "/xl/workbook.xml".into(),
    };
    assert_eq!(
        classify_xlsx_join_source(Some(&conflict)).code,
        Code::InvalidInput
    );

    let limit = litchi_xlsx::Error::Limit {
        kind: litchi_xlsx::LimitKind::ConditionalFormatting(
            litchi_xlsx::conditional_formatting::edit::LimitKind::Formats,
        ),
        observed: 2,
        maximum: 1,
    };
    assert_eq!(
        classify_xlsx_join_source(Some(&limit)).code,
        Code::LimitExceeded
    );

    let direct = litchi_xlsx::Error::SourceChanged {
        expected: SourceVersion::new(7, 1),
        observed: SourceVersion::new(7, 2),
    };
    let nested = litchi_xlsx::Error::Common(litchi_ooxml_common::Error::SourceChanged {
        expected: SourceVersion::new(7, 1),
        observed: SourceVersion::new(7, 2),
    });
    let direct_code = classify_xlsx_join_source(Some(&direct)).code;
    assert_eq!(direct_code, Code::IoError);
    assert_eq!(classify_xlsx_join_source(Some(&nested)).code, direct_code);
    assert_eq!(classify_xlsx_join_source(None).code, Code::CorruptFile);
}

#[cfg(feature = "docx")]
#[test]
fn docx_nested_metadata_refusal_and_source_change_keep_classification() {
    let refusal = litchi_core::edit::Refusal::new(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "set_custom_props",
        "rewriting unsupported custom-property content is refused",
    );
    assert_code(
        litchi_docx::metadata::Error::Properties(litchi_ooxml_common::properties::Error::Refused(
            refusal,
        )),
        Code::UnsupportedCapability,
    );
    let (expected, observed) = changed();
    assert_code(
        litchi_docx::metadata::Error::Properties(
            litchi_ooxml_common::properties::Error::SourceChanged { expected, observed },
        ),
        Code::IoError,
    );
}

#[cfg(feature = "docx")]
#[test]
fn docx_recorded_replay_conflicts_are_invalid_input() {
    use litchi_docx::Error as DocxError;

    assert_code(
        DocxError::PatchConflict {
            operation: "apply_section_layout",
            reason: "the patch source does not match the package",
        },
        Code::InvalidInput,
    );

    assert_code(
        DocxError::Opc(litchi_opc::OpcError::SourceArtifactMismatch {
            artifact: "current",
            field: "sha256",
        }),
        Code::InvalidInput,
    );
    assert_code(
        DocxError::Opc(litchi_opc::OpcError::SourceArtifactMismatch {
            artifact: "original",
            field: "length",
        }),
        Code::InvalidInput,
    );
    let (expected, actual) = changed();
    assert_code(
        DocxError::Opc(litchi_opc::OpcError::SourceChanged { expected, actual }),
        Code::IoError,
    );

    assert_code(
        DocxError::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "document_variables",
            "selected markup cannot be preserved",
        )),
        Code::UnsupportedCapability,
    );
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
/// The typed budgets of the shared OOXML families (Custom XML, Web
/// Extensions, document properties, Ribbon), carried by each OOXML crate's
/// shared-error variant, are limit errors (record 0360).
#[test]
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn shared_ooxml_family_limits_are_limit_exceeded() {
    let shared = || {
        [
            litchi_ooxml_common::Error::Limit {
                kind: litchi_ooxml_common::LimitKind::CustomXml(
                    litchi_ooxml_common::custom_xml::LimitKind::Items,
                ),
                observed: 11,
                maximum: 10,
            },
            litchi_ooxml_common::Error::Limit {
                kind: litchi_ooxml_common::LimitKind::Web(
                    litchi_ooxml_common::web::LimitKind::Panes,
                ),
                observed: 11,
                maximum: 10,
            },
            litchi_ooxml_common::Error::Limit {
                kind: litchi_ooxml_common::LimitKind::Ribbon(
                    litchi_ooxml_common::ribbon::LimitKind::XmlBytes,
                ),
                observed: 11,
                maximum: 10,
            },
            litchi_ooxml_common::Error::Limit {
                kind: litchi_ooxml_common::LimitKind::Properties(
                    litchi_ooxml_common::properties::LimitKind::RetainedBytes,
                ),
                observed: 11,
                maximum: 10,
            },
        ]
    };
    let check = |error: Error| assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");
    let mce = || litchi_ooxml_common::mce::Error::LimitExceeded {
        kind: litchi_ooxml_common::mce::LimitKind::EventBytes,
        observed: 17,
        maximum: 16,
    };
    #[cfg(feature = "docx")]
    {
        for error in shared() {
            check(litchi_docx::Error::Common(error).into());
        }
        check(litchi_docx::Error::Mce(mce()).into());
        check(litchi_docx::Error::Common(litchi_ooxml_common::Error::Mce(mce())).into());
    }
    #[cfg(feature = "pptx")]
    {
        for error in shared() {
            check(litchi_pptx::Error::CustomProperties(error).into());
        }
        check(litchi_pptx::Error::MarkupCompatibility(mce()).into());
        check(litchi_pptx::Error::CustomProperties(litchi_ooxml_common::Error::Mce(mce())).into());
    }
    #[cfg(feature = "xlsx")]
    {
        for error in shared() {
            check(litchi_xlsx::Error::Common(error).into());
        }
        check(litchi_xlsx::Error::MarkupCompatibility(mce()).into());
        check(litchi_xlsx::Error::Common(litchi_ooxml_common::Error::Mce(mce())).into());
    }
    #[cfg(feature = "xlsb")]
    {
        for error in shared() {
            check(litchi_xlsb::package::PackageError::Common(error).into());
        }
        check(
            litchi_xlsb::package::PackageError::Common(litchi_ooxml_common::Error::Mce(mce()))
                .into(),
        );
    }
}

#[cfg(feature = "docx")]
#[test]
fn docx_metadata_limit_conversion_keeps_details_and_classifies_downstream() {
    let cases = [
        (
            litchi_ooxml_common::properties::LimitKind::TextBytes,
            11,
            10,
        ),
        (litchi_ooxml_common::properties::LimitKind::XmlBytes, 13, 12),
        (
            litchi_ooxml_common::properties::LimitKind::XmlEvents,
            15,
            14,
        ),
        (
            litchi_ooxml_common::properties::LimitKind::CustomProperties,
            17,
            16,
        ),
        (litchi_ooxml_common::properties::LimitKind::Names, 19, 18),
        (
            litchi_ooxml_common::properties::LimitKind::RetainedBytes,
            21,
            20,
        ),
    ];
    for (kind, observed, maximum) in cases {
        let docx_error: litchi_docx::Error = litchi_ooxml_common::properties::Error::Limit {
            kind,
            observed,
            maximum,
        }
        .into();
        match &docx_error {
            litchi_docx::Error::Common(litchi_ooxml_common::Error::Limit {
                kind: common_kind,
                observed: actual,
                maximum: max,
            }) => {
                assert_eq!(
                    *common_kind,
                    litchi_ooxml_common::LimitKind::Properties(kind)
                );
                assert_eq!(*actual, observed);
                assert_eq!(*max, maximum);
                assert_eq!(common_kind.as_str(), kind.as_str());
            },
            other => panic!("properties limit lost its typed common wrapper: {other:?}"),
        }
        let unified: Error = litchi_docx::metadata::Error::Package(docx_error).into();
        assert_eq!(unified.code(), Code::LimitExceeded);
        assert!(std::error::Error::source(&unified).is_some());
    }
}

/// Shared semantic publication failures retain their contract when a format
/// adapter receives them through its common-error wrapper.
#[test]
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn shared_ooxml_semantic_errors_keep_their_wire_classes() {
    let shared = || {
        vec![
            (
                litchi_ooxml_common::Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "test",
                    "the operation is not supported",
                )),
                Code::UnsupportedCapability,
            ),
            (
                litchi_ooxml_common::Error::PatchConflict,
                Code::InvalidInput,
            ),
            (litchi_ooxml_common::Error::Signed, Code::DocumentProtected),
            (
                litchi_ooxml_common::Error::SourceChanged {
                    expected: SourceVersion::new(7, 1),
                    observed: SourceVersion::new(7, 2),
                },
                Code::IoError,
            ),
            (
                litchi_ooxml_common::Error::Internal("generated invariant"),
                Code::InternalError,
            ),
        ]
    };
    let check = |error: Error, expected| assert_eq!(error.code(), expected, "{error:?}");
    #[cfg(feature = "docx")]
    for (error, expected) in shared() {
        check(litchi_docx::Error::Common(error).into(), expected);
    }
    #[cfg(feature = "docx")]
    check(
        litchi_docx::Error::Common(litchi_ooxml_common::Error::Opc(
            litchi_opc::OpcError::SourceChanged {
                expected: SourceVersion::new(8, 1),
                actual: SourceVersion::new(8, 2),
            },
        ))
        .into(),
        Code::IoError,
    );
    #[cfg(feature = "pptx")]
    for (error, expected) in shared() {
        check(litchi_pptx::Error::CustomProperties(error).into(), expected);
    }
    #[cfg(feature = "xlsx")]
    for (error, expected) in shared() {
        check(litchi_xlsx::Error::Common(error).into(), expected);
    }
    #[cfg(feature = "xlsb")]
    for (error, expected) in shared() {
        check(
            litchi_xlsb::package::PackageError::Common(error).into(),
            expected,
        );
    }
}

/// A shared MCE policy rejects malformed caller bounds before reading input;
/// hosts must preserve that as an invalid argument rather than an exhausted file
/// resource (record 0360).
#[test]
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn shared_ooxml_invalid_mce_limits_are_invalid_arguments() {
    let check = |error: Error| assert_eq!(error.code(), Code::InvalidArgument, "{error:?}");
    let mce = || litchi_ooxml_common::mce::Error::InvalidLimit {
        kind: litchi_ooxml_common::mce::LimitKind::InputBytes,
        value: 0,
        minimum: 1,
        maximum: 2 * 1024 * 1024 * 1024,
    };
    #[cfg(feature = "docx")]
    {
        check(litchi_docx::Error::Mce(mce()).into());
        check(litchi_docx::Error::Common(litchi_ooxml_common::Error::Mce(mce())).into());
    }
    #[cfg(feature = "pptx")]
    {
        check(litchi_pptx::Error::MarkupCompatibility(mce()).into());
        check(litchi_pptx::Error::CustomProperties(litchi_ooxml_common::Error::Mce(mce())).into());
    }
    #[cfg(feature = "xlsx")]
    {
        check(litchi_xlsx::Error::MarkupCompatibility(mce()).into());
        check(litchi_xlsx::Error::Common(litchi_ooxml_common::Error::Mce(mce())).into());
    }
    #[cfg(feature = "xlsb")]
    {
        check(
            litchi_xlsb::package::PackageError::Common(litchi_ooxml_common::Error::Mce(mce()))
                .into(),
        );
    }
}

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
            kind: litchi_pptx::LimitKind::Codec(litchi_pptx::CodecLimitKind::Slides),
            observed: 11,
            maximum: 10,
        }
        .into(),
    );
    #[cfg(feature = "xlsx")]
    check(litchi_xlsx::Error::ResourceLimit(resource_limit()).into());
    #[cfg(feature = "xlsb")]
    {
        check(
            litchi_xlsb::package::PackageError::Limit {
                kind: litchi_xlsb::cell_values::LimitKind::Cells.into(),
                observed: 11,
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
        litchi_xlsb::package::PackageError::UnsupportedFeature("decrypt"),
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
    #[cfg(feature = "rtf")]
    assert_code(
        litchi_rtf::Error::UnsupportedContent("multiple equations in one math paragraph"),
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
    #[cfg(feature = "pptx")]
    assert_code(
        litchi_pptx::Error::CapacityOverflow {
            resource: "replay content fingerprint",
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
fn xlsx_refusals_conflicts_and_limits_have_their_codes() {
    use litchi_core::edit::{Refusal, RefusalKind};
    assert_code(
        litchi_xlsx::Error::Refused(Refusal::new(
            RefusalKind::NotFound,
            "edit_auto_filter",
            "the worksheet does not exist",
        )),
        Code::NotFound,
    );
    assert_code(
        litchi_xlsx::Error::Refused(Refusal::new(
            RefusalKind::UnsupportedContent,
            "edit_comments",
            "markup compatibility",
        )),
        Code::UnsupportedCapability,
    );
    assert_code(
        litchi_xlsx::Error::PatchConflict {
            part: "/xl/workbook.xml".into(),
        },
        Code::InvalidInput,
    );
    assert_code(
        litchi_xlsx::Error::InvalidLimit {
            kind: litchi_xlsx::LimitKind::Opc(litchi_xlsx::opc::LimitKind::Parts),
            value: 0,
            maximum: 1,
        },
        Code::InvalidArgument,
    );
    assert_code(
        litchi_xlsx::Error::Limit {
            kind: litchi_xlsx::LimitKind::ConditionalFormatting(
                litchi_xlsx::conditional_formatting::edit::LimitKind::Formats,
            ),
            observed: 2,
            maximum: 1,
        },
        Code::LimitExceeded,
    );
    assert_code(
        litchi_xlsx::Error::Limit {
            kind: litchi_xlsx::LimitKind::Connections(
                litchi_xlsx::connections::LimitKind::NamespaceBindings,
            ),
            observed: 5,
            maximum: 4,
        },
        Code::LimitExceeded,
    );
    assert_code(
        litchi_xlsx::Error::FormControl(litchi_xlsx::form_control::FormControlError::Limit {
            kind: litchi_xlsx::form_control::LimitKind::Items,
            observed: 5,
            maximum: 4,
        }),
        Code::LimitExceeded,
    );
}

#[cfg(feature = "xlsb")]
#[test]
fn xlsb_refusals_conflicts_and_invalid_limits_classify_by_kind() {
    use litchi_core::edit::{Refusal, RefusalKind};
    use litchi_xlsb::package::PackageError;

    assert_code(
        PackageError::Refused(Refusal::new(
            RefusalKind::UnsupportedContent,
            "edit_timelines",
            "non-canonical source",
        )),
        Code::UnsupportedCapability,
    );
    assert_code(
        PackageError::PatchConflict {
            part: "Custom Data source closure",
        },
        Code::InvalidInput,
    );
    assert_code(
        PackageError::InvalidLimit {
            kind: litchi_xlsb::package::LimitKind::Opc(litchi_xlsb::opc::LimitKind::Parts),
            value: 0,
            maximum: 1,
        },
        Code::InvalidArgument,
    );
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
    use litchi_formula::omml::LimitKind;
    use litchi_formula::{FormulaError, LatexError, OmmlError};

    let unsupported = assert_code(
        FormulaError::Omml(OmmlError::UnsupportedFeature("m:groupChr".into())),
        Code::UnsupportedElement,
    );
    assert!(matches!(unsupported.leaf(), Source::Formula(_)));
    assert_code(
        FormulaError::Omml(OmmlError::LimitExceeded {
            kind: LimitKind::Depth,
            limit: 64,
            actual: 65,
        }),
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
    assert_code(EditError::PatchConflict, Code::InvalidInput);
    // A three-way semantic overlap remains a distinct body conflict.
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

#[cfg(feature = "doc")]
#[test]
fn legacy_doc_foreign_boundaries_delegate_classification_before_core_projection() {
    assert_code(
        litchi_doc::Error::Signature(litchi_sign::Error::Limit("secret-limit".into())),
        Code::LimitExceeded,
    );
    assert_code(
        litchi_doc::Error::Signature(litchi_sign::Error::Cfb(litchi_cfb::OleError::Io(
            io::Error::other("secret-io"),
        ))),
        Code::IoError,
    );
    assert_code(
        litchi_doc::Error::Signature(litchi_sign::Error::Unsupported("secret-unsupported".into())),
        Code::UnsupportedCapability,
    );
    assert_code(
        litchi_doc::Error::Vba(litchi_vba::Error::Limit {
            kind: litchi_vba::LimitKind::Modules,
            observed: 2,
            maximum: 1,
        }),
        Code::LimitExceeded,
    );
    assert_code(
        litchi_doc::Error::Crypto(litchi_crypto::spaces::Error::PatchConflict),
        Code::InvalidInput,
    );
    assert_code(
        litchi_doc::Error::Crypto(litchi_crypto::spaces::Error::InvalidLimit {
            kind: litchi_crypto::spaces::LimitKind::OutputBytes,
            value: 0,
            maximum: 1,
        }),
        Code::InvalidArgument,
    );
    assert_code(
        litchi_doc::Error::Crypto(litchi_crypto::spaces::Error::LimitExceeded {
            kind: litchi_crypto::spaces::LimitKind::TotalBytes,
            observed: 2,
            maximum: 1,
        }),
        Code::LimitExceeded,
    );
}

#[cfg(feature = "xls")]
#[test]
fn legacy_xls_foreign_boundaries_delegate_classification() {
    assert_code(
        litchi_xls::Error::Signature(litchi_sign::Error::Limit("secret-limit".into())),
        Code::LimitExceeded,
    );
    assert_code(
        litchi_xls::Error::Vba(litchi_vba::Error::UnsupportedCodePage(65_535)),
        Code::UnsupportedCapability,
    );
    assert_code(
        litchi_xls::Error::CustomXml(litchi_ole_common::custom_xml::Error::ResourceLimit(
            "secret-limit".into(),
        )),
        Code::LimitExceeded,
    );
}

#[cfg(feature = "ppt")]
#[test]
fn legacy_ppt_custom_xml_boundary_delegates_classification() {
    assert_code(
        litchi_ppt::Error::CustomXml(litchi_ole_common::custom_xml::Error::ResourceLimit(
            "secret-limit".into(),
        )),
        Code::LimitExceeded,
    );
}

#[cfg(all(feature = "ppt", feature = "sign"))]
#[test]
fn legacy_ppt_signature_boundary_delegates_classification() {
    assert_code(
        litchi_ppt::Error::Signature(litchi_sign::Error::Limit("secret-limit".into())),
        Code::LimitExceeded,
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

#[cfg(any(feature = "docx", feature = "pptx"))]
#[test]
fn drawingml_typed_errors_keep_host_classification() {
    #[cfg(feature = "docx")]
    {
        let limit = assert_code(
            litchi_docx::Error::Drawing(litchi_drawingml::Error::Limit {
                kind: litchi_drawingml::LimitKind::ChartTextBytes,
                observed: 9,
                maximum: 8,
            }),
            Code::LimitExceeded,
        );
        assert!(matches!(limit.leaf(), Source::Docx(_)));
        assert_code(
            litchi_docx::Error::Drawing(litchi_drawingml::Error::InvalidLimit {
                kind: litchi_drawingml::LimitKind::ChartTextBytes,
                value: 0,
                maximum: 8,
            }),
            Code::InvalidArgument,
        );
        assert_code(
            litchi_docx::Error::Drawing(litchi_drawingml::Error::PatchConflict),
            Code::InvalidInput,
        );
    }

    #[cfg(feature = "pptx")]
    {
        let limit = assert_code(
            litchi_pptx::Error::Drawing(litchi_drawingml::Error::Limit {
                kind: litchi_drawingml::LimitKind::ChartTextBytes,
                observed: 9,
                maximum: 8,
            }),
            Code::LimitExceeded,
        );
        assert!(matches!(limit.leaf(), Source::Pptx(_)));
        assert_code(
            litchi_pptx::Error::Drawing(litchi_drawingml::Error::InvalidLimit {
                kind: litchi_drawingml::LimitKind::ChartTextBytes,
                value: 0,
                maximum: 8,
            }),
            Code::InvalidArgument,
        );
        assert_code(
            litchi_pptx::Error::Drawing(litchi_drawingml::Error::PatchConflict),
            Code::InvalidInput,
        );
    }
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
    let error = Error::from(litchi_xlsb::package::PackageError::InvalidFormat(secret));
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
        (Refusal::FinalSlide, Code::InvalidValue),
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
        let reason = match kind {
            Refusal::SignedPackage => "the package is signed",
            Refusal::FinalSlide => "a presentation must retain at least one slide",
            _ => "the presentation cannot satisfy the requested slide removal",
        };
        assert_code(
            litchi_pptx::Error::Refused(litchi_core::edit::Refusal::new(
                kind.kind(),
                "slide_removal",
                reason,
            )),
            code,
        );
    }
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_refusals_classify_by_kind_not_operation_or_reason() {
    use litchi_core::edit::{Refusal, RefusalKind};

    let expected = |kind| match kind {
        RefusalKind::UnsupportedOperation | RefusalKind::UnsupportedContent => {
            Code::UnsupportedCapability
        },
        RefusalKind::UnsupportedTarget => Code::UnsupportedType,
        RefusalKind::UnsupportedProperty => Code::UnsupportedProperty,
        RefusalKind::InvalidValue => Code::InvalidValue,
        RefusalKind::NotFound => Code::NotFound,
        RefusalKind::Protected => Code::DocumentProtected,
        _ => Code::UnsupportedCapability,
    };

    for kind in RefusalKind::ALL.iter().copied() {
        for (operation, reason) in [
            ("slide_removal", "the package is signed"),
            ("unrelated_operation", "caller-authored reason marker"),
        ] {
            let error = assert_code(
                litchi_pptx::Error::Refused(Refusal::new(kind, operation, reason)),
                expected(kind),
            );
            assert!(
                !error.to_string().contains("caller-authored reason marker"),
                "refusal diagnostics must remain content-free: {error}"
            );
        }
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
        let native = if matches!(kind, Refusal::AmbiguousBindings) {
            litchi_pptx::Error::Invalid("slide-order bindings are ambiguous".into())
        } else {
            litchi_pptx::Error::Refused(litchi_core::edit::Refusal::new(
                kind.kind(),
                "slide_order",
                "the presentation contains unsupported positional metadata",
            ))
        };
        assert_code(native, code);
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
            Native::Body(litchi_doc::body_text::Error::PatchConflict),
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
fn ppt_officeart_limits_are_not_reported_as_corrupt_files() {
    use litchi_ppt::odraw::{Error as OfficeArtError, ImageLimitKind, LimitKind};

    for (officeart, expected) in [
        (
            OfficeArtError::LimitExceeded {
                limit: LimitKind::Records,
                observed: 9,
                maximum: 8,
            },
            Code::LimitExceeded,
        ),
        (
            OfficeArtError::ImageLimitExceeded {
                limit: ImageLimitKind::Files,
                observed: 9,
                maximum: 8,
            },
            Code::LimitExceeded,
        ),
        (
            OfficeArtError::InvalidLimit {
                limit: LimitKind::Depth,
                value: 65,
                maximum: 64,
            },
            Code::InvalidArgument,
        ),
        (
            OfficeArtError::MalformedShape {
                reason: "shape invariant",
            },
            Code::CorruptFile,
        ),
    ] {
        let error = assert_code(litchi_ppt::Error::OfficeArt(officeart), expected);
        assert!(matches!(
            error.leaf(),
            Source::Ppt(litchi_ppt::Error::OfficeArt(_))
        ));
    }
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

#[cfg(feature = "odp")]
#[test]
fn odp_source_append_public_errors_classify_stale_and_patch_bases() {
    use litchi_core::SourceVersion;
    use litchi_odf_common::SourceContentPublicationProgress;
    use litchi_odp::tail_append::Error as AppendError;

    let stale = Error::from(AppendError::SourceChanged {
        expected: SourceVersion::new(0x100, 1),
        observed: SourceVersion::new(0x100, 2),
        progress: SourceContentPublicationProgress::Complete { bytes: 17 },
        source: None,
    });
    assert_eq!(stale.code(), Code::IoError);
    assert!(matches!(
        stale.leaf(),
        Source::OdpTailAppend(AppendError::SourceChanged { .. })
    ));

    let conflict = Error::from(AppendError::PatchConflict {
        operation: "odp.tail_append.apply",
        reason: "recorded base differs",
    });
    assert_eq!(conflict.code(), Code::InvalidInput);
    assert!(matches!(
        conflict.leaf(),
        Source::OdpTailAppend(AppendError::PatchConflict { .. })
    ));
}

#[cfg(feature = "odp")]
#[test]
fn odp_source_publication_errors_keep_typed_progress_through_facade() {
    use super::Classify as _;
    use litchi_core::SourceVersion;
    use litchi_odf_common::{
        Error as OdfError, SourceContentPublicationError, SourceContentPublicationProgress,
    };

    let expected = SourceVersion::new(0x200, 1);
    let observed = SourceVersion::new(0x200, 2);
    let source_error = OdfError::from(SourceContentPublicationError::SourceChanged {
        expected,
        observed,
        progress: SourceContentPublicationProgress::Complete { bytes: 23 },
    });
    assert_eq!(source_error.classify().code, Code::IoError);
    let unified = Error::from(source_error);
    assert_eq!(unified.code(), Code::IoError);
    assert!(matches!(
        unified.leaf(),
        Source::Odf(OdfError::SourceContentPublication(_))
    ));
}

#[cfg(feature = "odp")]
#[test]
fn odp_archive_replay_prefers_typed_stale_callback_over_content_fault() {
    use litchi_core::SourceVersion;
    use litchi_odf_common::{
        Error as OdfError, SourceContentInsertionCallbackError, SourceContentInsertionError,
        SourceContentPublicationError, SourceContentPublicationProgress, SourceMemberReaderError,
    };
    use soapberry_zip::{ErrorKind, ReplayPass, ReplayProgress, ReplayPublicationError};

    let expected = SourceVersion::new(0x300, 1);
    let observed = SourceVersion::new(0x300, 2);
    let stale_callback = || {
        SourceContentInsertionCallbackError::Member(Box::new(SourceMemberReaderError::Core {
            error: litchi_core::Error::SourceChanged { expected, observed },
            callback_error: Some(SourceContentInsertionCallbackError::Verification {
                what: "secondary proof",
            }),
        }))
    };
    let replay = ReplayPublicationError::Archive {
        pass: ReplayPass::Emit,
        source: soapberry_zip::Error::from(ErrorKind::InvalidInput {
            msg: "malformed archive member".to_string(),
        }),
        callback_error: Some(stale_callback()),
        progress: ReplayProgress::Complete { bytes: 19 },
    };
    let odf = OdfError::from(SourceContentInsertionError::Replay {
        progress: SourceContentPublicationProgress::Complete { bytes: 19 },
        source: replay,
    });

    let unified = Error::from(odf);
    assert_eq!(unified.code(), Code::IoError);
    assert!(matches!(
        unified.leaf(),
        Source::Odf(OdfError::SourceContentInsertion(_))
    ));

    let archive_with_cancelled_callback = ReplayPublicationError::Archive {
        pass: ReplayPass::Emit,
        source: soapberry_zip::Error::from(ErrorKind::InvalidInput {
            msg: "malformed archive member".to_string(),
        }),
        callback_error: Some(SourceContentInsertionCallbackError::Cancelled),
        progress: ReplayProgress::Untouched,
    };
    let cancelled = Error::from(OdfError::from(SourceContentInsertionError::Replay {
        progress: SourceContentPublicationProgress::Untouched,
        source: archive_with_cancelled_callback,
    }));
    assert_eq!(cancelled.code(), Code::CorruptFile);

    let transport = Error::from(OdfError::from(SourceContentInsertionError::Transport {
        progress: SourceContentPublicationProgress::Complete { bytes: 19 },
        publication: SourceContentPublicationError::Unsupported {
            reason: "unsupported test source".to_string(),
        },
        replay: Box::new(ReplayPublicationError::Callback {
            pass: ReplayPass::Emit,
            source: stale_callback(),
            progress: ReplayProgress::Complete { bytes: 19 },
        }),
    }));
    assert_eq!(transport.code(), Code::IoError);
}

#[cfg(feature = "odp")]
#[test]
fn odp_public_source_replay_classifies_actual_stale_and_conflicting_inputs() {
    use std::io;
    use std::sync::atomic::{AtomicU64, Ordering};

    use litchi_core::{OwnedSource, ReadAt};
    use litchi_odf_common::constants::{ODF_CONTENT, ODF_PRESENTATION};
    use litchi_odf_common::core::PackageWriter;
    use litchi_odp::SourceBackedPresentation;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

    fn package_bytes(page_name: &str) -> Vec<u8> {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}"><office:body><office:presentation><draw:page draw:name="{page_name}"/></office:presentation></office:body></office:document-content>"#
        );
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer.set_mimetype(ODF_PRESENTATION).unwrap();
        writer.add_file(ODF_CONTENT, content.as_bytes()).unwrap();
        writer.finish_to_bytes().unwrap()
    }

    struct VersionedSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
    }

    impl ReadAt for VersionedSource {
        fn len(&self) -> io::Result<u64> {
            u64::try_from(self.bytes.len())
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "length overflow"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                0x4f44_505f_6661_6365,
                self.revision.load(Ordering::Relaxed),
            ))
        }
    }

    let source = Arc::new(VersionedSource {
        bytes: package_bytes("page1"),
        revision: AtomicU64::new(0),
    });
    let owner = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let commit = owner
        .edit_tail_append("title", "body")
        .unwrap()
        .commit()
        .unwrap();

    source.revision.fetch_add(1, Ordering::Relaxed);
    let stale_leaf = match owner.apply_tail_append(commit.patch()) {
        Ok(_) => panic!("a changed source must reject the replay"),
        Err(error) => error,
    };
    let stale = Error::from(stale_leaf);
    assert_eq!(stale.code(), Code::IoError);
    assert!(matches!(stale.leaf(), Source::OdpTailAppend(_)));

    let conflicting = SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(
        package_bytes("other-page"),
    )))
    .unwrap();
    let conflict_leaf = match conflicting.apply_tail_append(commit.patch()) {
        Ok(_) => panic!("a different recipient must reject the patch"),
        Err(error) => error,
    };
    let conflict = Error::from(conflict_leaf);
    assert_eq!(conflict.code(), Code::InvalidInput);
    assert!(matches!(conflict.leaf(), Source::OdpTailAppend(_)));
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
    for operation in [
        "external_hyperlink_detachment",
        "external_hyperlink_redaction",
        "story_hyperlink_redaction",
        "document_variables",
    ] {
        assert_code(
            litchi_docx::Error::PatchConflict {
                operation,
                reason: "the patch base does not match the supplied source",
            },
            Code::InvalidInput,
        );
    }
    let producer = litchi_opc::ArtifactStreamError::Producer(litchi_docx::Error::PatchConflict {
        operation: "stream",
        reason: "another base",
    });
    assert_code(
        litchi_docx::Error::ArtifactStream(Box::new(producer)),
        Code::InvalidInput,
    );
    let partial = litchi_opc::ArtifactStreamError::Publication {
        written: 1,
        source: Box::new(litchi_opc::ArtifactStreamError::Producer(
            litchi_docx::Error::PatchConflict {
                operation: "stream",
                reason: "another base",
            },
        )),
        secondary: None,
    };
    assert_code(
        litchi_docx::Error::ArtifactStream(Box::new(partial)),
        Code::IoError,
    );
    assert_code(
        litchi_docx::Error::Internal {
            operation: "edit",
            reason: "invariant",
        },
        Code::InternalError,
    );
    assert_code(
        litchi_docx::Error::TailAppend(Box::new(
            litchi_docx::source_backed::tail_append::Error::Execution(ExecutionError::Cancelled),
        )),
        Code::Cancelled,
    );
    assert_code(
        litchi_docx::Error::TailAppend(Box::new(
            litchi_docx::source_backed::tail_append::Error::PatchConflict,
        )),
        Code::InvalidInput,
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

#[cfg(any(feature = "doc", feature = "xls", feature = "ppt"))]
#[test]
fn shared_object_errors_keep_codes_through_host_and_core_conversions() {
    use litchi_ole_common::{Error as ObjectError, LimitKind};
    let cases: [(fn() -> ObjectError, Code); 5] = [
        (
            || {
                ObjectError::Cfb(litchi_cfb::OleError::LimitExceeded {
                    resource: "input bytes",
                    observed: 2,
                    maximum: 1,
                })
            },
            Code::LimitExceeded,
        ),
        (
            || {
                ObjectError::Cfb(litchi_cfb::OleError::InvalidLimit {
                    resource: "input bytes",
                    value: 0,
                    maximum: 1,
                })
            },
            Code::InvalidArgument,
        ),
        (
            || ObjectError::Limit {
                kind: LimitKind::StreamBytes,
                observed: 2,
                maximum: 1,
            },
            Code::LimitExceeded,
        ),
        (
            || ObjectError::InvalidLimit {
                kind: LimitKind::StorageDepth,
                value: 0,
                maximum: 32,
            },
            Code::InvalidArgument,
        ),
        (|| ObjectError::PatchConflict, Code::InvalidInput),
    ];
    for (make_error, expected) in cases {
        let core: litchi_core::Error = make_error().into();
        assert_eq!(Error::from(core).code(), expected);
        #[cfg(feature = "doc")]
        {
            let host: litchi_doc::Error = make_error().into();
            assert_eq!(Error::from(host).code(), expected);
        }
        #[cfg(feature = "xls")]
        {
            let host: litchi_xls::Error = make_error().into();
            assert_eq!(Error::from(host).code(), expected);
        }
    }
    let changed = || {
        ObjectError::Cfb(litchi_cfb::OleError::SourceChanged {
            expected: SourceVersion::new(1, 0),
            observed: SourceVersion::new(1, 1),
        })
    };
    let core: litchi_core::Error = changed().into();
    assert_eq!(Error::from(core).code(), Code::IoError);
    #[cfg(feature = "doc")]
    assert_eq!(
        Error::from(litchi_doc::Error::from(changed())).code(),
        Code::IoError
    );
    #[cfg(feature = "xls")]
    assert_eq!(
        Error::from(litchi_xls::Error::from(changed())).code(),
        Code::IoError
    );
}

#[cfg(feature = "doc")]
#[test]
fn legacy_raw_replay_categories_reach_the_facade_unchanged() {
    // Raw OLE/property-set parse failures remain malformed input.  A stale
    // package patch uses the shared typed conflict and is caller input, while
    // an actual positional source change remains an I/O condition.
    assert_code(
        litchi_cfb::OleError::InvalidFormat("malformed property set".into()),
        Code::CorruptFile,
    );
    assert_code(
        litchi_doc::Error::Ole(litchi_cfb::OleError::InvalidFormat(
            "malformed property set".into(),
        )),
        Code::CorruptFile,
    );
    assert_code(
        litchi_doc::Error::from(litchi_ole_common::Error::PatchConflict),
        Code::InvalidInput,
    );
    assert_code(litchi_doc::Error::PatchConflict, Code::InvalidInput);
    assert_code(
        litchi_doc::Error::from(litchi_ole_common::Error::Cfb(
            litchi_cfb::OleError::SourceChanged {
                expected: SourceVersion::new(3, 1),
                observed: SourceVersion::new(3, 2),
            },
        )),
        Code::IoError,
    );
}

#[cfg(feature = "doc")]
#[test]
fn public_doc_stale_replay_reaches_the_facade_as_invalid_input() {
    let mut builder = litchi_doc::Builder::new();
    builder.add_paragraph("before").expect("fixture paragraph");
    let source = builder.build().expect("fixture document");

    let mut first = source.edit().expect("first edit");
    first
        .replace_paragraph(litchi_core::Position::new(0), "first")
        .expect("first replacement");
    let patch = first.commit().expect("first commit").patch().clone();

    let mut second = source.edit().expect("second edit");
    second
        .replace_paragraph(litchi_core::Position::new(0), "second")
        .expect("second replacement");
    let changed = second.commit().expect("second commit").into_package();
    let native = changed.apply(&patch).expect_err("stale package replay");
    assert!(matches!(
        &native,
        litchi_doc::package::edit::Error::Body(litchi_doc::body_text::Error::PatchConflict)
    ));

    let facade = Error::from(native);
    assert_eq!(facade.code(), Code::InvalidInput);
    assert!(matches!(facade.leaf(), Source::DocTransaction(_)));
}
