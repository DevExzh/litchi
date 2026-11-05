//! Contextual OOXML ingestion limits through the unified opener.
//!
//! The OPC read limits of DOCX, PPTX, XLSX and XLSB are one
//! [`litchi::OpenOptions::ooxml_limits`] knob, which [`litchi::open_with`]
//! and [`litchi::open_bytes_with`] apply on the path and byte routes. A
//! refused input reports [`litchi::Code::LimitExceeded`] and keeps the
//! limit, the observed value and the resource in its typed OPC source.
//!
//! M0a step S8c ported these tests from the facade's `Document`,
//! `Presentation` and `sheet` openers with a `*_with_limits` form.

#![cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]
#![cfg_attr(
    not(all(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")),
    allow(dead_code, reason = "helpers are shared by feature-gated tests")
)]

use std::path::{Path, PathBuf};

use litchi::error::Source;
use litchi::opc::{OpcError, SourceLimitKind, SourceLimits};
use litchi::{Code, File, OpenOptions};

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// The unified openers that take the limits, by type.
fn assert_openers_are_nameable() {
    let _: fn(&Path, &OpenOptions) -> litchi::Result<File> =
        |path, options| litchi::open_with(path, options);
    let _: fn(Vec<u8>, &OpenOptions) -> litchi::Result<File> =
        |bytes, options| litchi::open_bytes_with(bytes, options);
    let _: fn(
        Vec<u8>,
        &litchi::common::detection::SourceLimits,
    ) -> Option<litchi::common::detection::DetectedFormat> =
        litchi::common::detection::detect_format_smart;
}

/// Limits whose input ceiling is one byte below `bytes`.
fn one_byte_short(bytes: &[u8]) -> SourceLimits {
    SourceLimits::builder()
        .max_input_bytes(u64::try_from(bytes.len() - 1).expect("fixture length fits u64"))
        .expect("positive input limit")
        .build()
        .expect("valid limits")
}

/// `error` is the OPC input-size refusal: `observed` bytes against a ceiling
/// of `maximum`.
#[track_caller]
fn assert_input_limit(error: &litchi::Error, observed: u64, maximum: u64) {
    assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");
    assert!(
        matches!(
            error.leaf(),
            Source::Opc(OpcError::Limit {
                kind: SourceLimitKind::InputBytes,
                observed: actual,
                maximum: limit,
            }) if *actual == observed && *limit == maximum
        ),
        "expected the {maximum}-byte input limit to refuse {observed} bytes: {error:?}"
    );
}

fn len(bytes: &[u8]) -> u64 {
    u64::try_from(bytes.len()).expect("fixture length fits u64")
}

// ---------------------------------------------------------------- documents --

#[cfg(feature = "docx")]
#[test]
fn docx_limits_are_nameable() {
    let limits: litchi::docx::SourceLimits = Default::default();
    // The DOCX limits are the OPC limits the unified opener takes.
    let _ = OpenOptions::new().ooxml_limits(limits);
    assert_openers_are_nameable();
}

#[cfg(feature = "docx")]
#[test]
fn document_limits_bound_path_and_byte_ingress() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("bounded.docx");
    let package = litchi::docx::Package::new().expect("new DOCX package");
    package.save(&path).expect("save DOCX package");

    let bytes = std::fs::read(&path).expect("read DOCX fixture");
    let options = OpenOptions::new().ooxml_limits(one_byte_short(&bytes));

    let file = litchi::open(&path).expect("the default limits admit the package");
    assert_eq!(file.format(), Some(litchi::FileFormat::Docx));
    let observed = len(&bytes);
    assert_input_limit(
        &litchi::open_with(&path, &options).unwrap_err(),
        observed,
        observed - 1,
    );
    assert_input_limit(
        &litchi::open_bytes_with(bytes, &options).unwrap_err(),
        observed,
        observed - 1,
    );
}

#[cfg(feature = "docx")]
#[test]
fn document_limits_reject_invalid_suffix_and_unknown_extension_ooxml() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let invalid = directory.path().join("oversized.docx");
    std::fs::write(&invalid, vec![b'x'; 32]).expect("write invalid DOCX");
    let limits = SourceLimits::builder()
        .max_input_bytes(8)
        .expect("positive input limit")
        .build()
        .expect("valid limits");
    assert_input_limit(
        &litchi::open_with(&invalid, &OpenOptions::new().ooxml_limits(limits)).unwrap_err(),
        32,
        8,
    );

    let unknown = directory.path().join("bounded-package");
    let package = litchi::docx::Package::new().expect("new DOCX package");
    package.save(&unknown).expect("save DOCX package");
    let bytes = std::fs::read(&unknown).expect("read DOCX fixture");
    let tight = one_byte_short(&bytes);
    assert_input_limit(
        &litchi::open_with(&unknown, &OpenOptions::new().ooxml_limits(tight)).unwrap_err(),
        len(&bytes),
        len(&bytes) - 1,
    );
    assert!(litchi::detection_smart::ooxml::detect_zip_format(&bytes, &tight).is_none());
}

// ------------------------------------------------------------ presentations --

#[cfg(feature = "pptx")]
#[test]
fn pptx_limits_are_nameable() {
    let limits: litchi::pptx::SourceLimits = Default::default();
    let _ = OpenOptions::new().ooxml_limits(limits);
    assert_openers_are_nameable();
}

#[cfg(feature = "pptx")]
#[test]
fn presentation_limits_bound_path_and_byte_ingress() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("bounded.pptx");
    let package = litchi::pptx::Package::new().expect("new PPTX package");
    package.save(&path).expect("save PPTX package");

    let bytes = std::fs::read(&path).expect("read PPTX fixture");
    let options = OpenOptions::new().ooxml_limits(one_byte_short(&bytes));

    let file = litchi::open(&path).expect("the default limits admit the package");
    assert_eq!(file.format(), Some(litchi::FileFormat::Pptx));
    let observed = len(&bytes);
    assert_input_limit(
        &litchi::open_with(&path, &options).unwrap_err(),
        observed,
        observed - 1,
    );
    assert_input_limit(
        &litchi::open_bytes_with(bytes, &options).unwrap_err(),
        observed,
        observed - 1,
    );
}

#[cfg(feature = "pptx")]
#[test]
fn presentation_limits_reject_invalid_suffix_and_unknown_extension_ooxml() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let invalid = directory.path().join("oversized.pptx");
    std::fs::write(&invalid, vec![b'x'; 32]).expect("write invalid PPTX");
    let limits = SourceLimits::builder()
        .max_input_bytes(8)
        .expect("positive input limit")
        .build()
        .expect("valid limits");
    assert_input_limit(
        &litchi::open_with(&invalid, &OpenOptions::new().ooxml_limits(limits)).unwrap_err(),
        32,
        8,
    );

    let unknown = directory.path().join("bounded-package");
    let package = litchi::pptx::Package::new().expect("new PPTX package");
    package.save(&unknown).expect("save PPTX package");
    let bytes = std::fs::read(&unknown).expect("read PPTX fixture");
    let tight = one_byte_short(&bytes);
    assert_input_limit(
        &litchi::open_with(&unknown, &OpenOptions::new().ooxml_limits(tight)).unwrap_err(),
        len(&bytes),
        len(&bytes) - 1,
    );
    assert!(litchi::detection_smart::ooxml::detect_zip_format(&bytes, &tight).is_none());
}

// ---------------------------------------------------------------- workbooks --

/// The old test named the facade's bounded XLSX sheet openers; the unified
/// openers take the same limits, and a workbook one byte over them is
/// refused on both routes.
#[cfg(feature = "xlsx")]
#[test]
fn xlsx_limits_and_bounded_sheet_helpers_are_nameable() {
    let limits: litchi::xlsx::SourceLimits = Default::default();
    let _ = OpenOptions::new().ooxml_limits(limits);
    assert_openers_are_nameable();

    let path = fixture("ooxml/xlsx/SimpleNormal.xlsx");
    let bytes = std::fs::read(&path).unwrap();
    let options = OpenOptions::new().ooxml_limits(one_byte_short(&bytes));
    assert_eq!(
        litchi::open_bytes(bytes.clone()).unwrap().format(),
        Some(litchi::FileFormat::Xlsx)
    );
    let observed = len(&bytes);
    assert_input_limit(
        &litchi::open_with(&path, &options).unwrap_err(),
        observed,
        observed - 1,
    );
    assert_input_limit(
        &litchi::open_bytes_with(bytes, &options).unwrap_err(),
        observed,
        observed - 1,
    );
}

/// As for XLSX: the facade's bounded XLSB openers (eager, trait object, path
/// and bytes) are the unified openers with the same limits.
#[cfg(feature = "xlsb")]
#[test]
fn xlsb_limits_and_bounded_sheet_helpers_are_nameable() {
    let limits: litchi::xlsb::SourceLimits = Default::default();
    let _ = OpenOptions::new().ooxml_limits(limits);
    assert_openers_are_nameable();

    let path = fixture("ooxml/xlsb/Simple.xlsb");
    let bytes = std::fs::read(&path).unwrap();
    let options = OpenOptions::new().ooxml_limits(one_byte_short(&bytes));
    let eager = options.clone().source_mode(litchi::SourceMode::Eager);
    assert_eq!(
        litchi::open_bytes(bytes.clone()).unwrap().format(),
        Some(litchi::FileFormat::Xlsb)
    );
    let observed = len(&bytes);
    for error in [
        litchi::open_with(&path, &options).unwrap_err(),
        litchi::open_with(&path, &eager).unwrap_err(),
        litchi::open_bytes_with(bytes.clone(), &options).unwrap_err(),
        litchi::open_bytes_with(bytes, &eager).unwrap_err(),
    ] {
        assert_input_limit(&error, observed, observed - 1);
    }
}
