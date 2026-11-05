//! A wrong-family read reports which format was detected.
//!
//! Detection classifies the input before a family view refuses it, so
//! refusing with a bare "not an Office file" throws away something the caller
//! needs. These tests pin the detected format onto each family view's
//! refusal ([`litchi::File::document`], [`litchi::File::presentation`] and
//! [`litchi::File::workbook`] return [`Code::UnsupportedFormat`] with
//! [`litchi::Error::detected`] naming the format), and keep the genuinely
//! unrecognized case reporting no detected format.
//!
//! M0a step S8c ported these tests from the facade's format-family openers
//! (document, presentation and workbook), which refused the same inputs with
//! `UnexpectedFormat { detected }` and `NotOfficeFile`.

#![cfg(all(
    any(unix, windows),
    any(feature = "docx", feature = "pptx", feature = "xlsx")
))]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]
#![cfg_attr(
    not(all(
        feature = "doc",
        feature = "docx",
        feature = "odt",
        feature = "ppt",
        feature = "pptx",
        feature = "odp",
        feature = "xlsx"
    )),
    allow(dead_code, reason = "helpers are shared by feature-gated tests")
)]

use std::path::{Path, PathBuf};

use litchi::{Code, Error, FileFormat};

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// Open `relative`, which must open, for a family view to refuse.
fn open(relative: &str) -> litchi::File {
    litchi::open(fixture(relative)).unwrap_or_else(|error| panic!("{relative}: {error:?}"))
}

#[track_caller]
fn assert_unexpected(error: &Error, expected: FileFormat) {
    assert_eq!(error.code(), Code::UnsupportedFormat, "{error:?}");
    assert_eq!(error.detected(), Some(expected), "{error:?}");
}

/// The unified opener's refusal of an input no format claims: the old
/// openers' `NotOfficeFile`.
#[track_caller]
fn assert_not_an_office_file(relative: &str) {
    let error = litchi::open(fixture(relative))
        .err()
        .unwrap_or_else(|| panic!("{relative} is not an Office file"));
    assert_eq!(error.code(), Code::UnsupportedFormat, "{error:?}");
    assert_eq!(error.detected(), None, "{error:?}");
}

// ---------------------------------------------------------------- documents --

#[cfg(all(feature = "docx", feature = "xlsx"))]
#[test]
fn an_xlsx_given_to_the_document_opener_reports_xlsx() {
    let error = open("ooxml/xlsx/styles.xlsx")
        .document()
        .expect_err("an XLSX is not a document");
    assert_unexpected(&error, FileFormat::Xlsx);
}

#[cfg(all(feature = "docx", feature = "pptx"))]
#[test]
fn a_pptx_given_to_the_document_opener_reports_pptx() {
    let error = open("ooxml/pptx/shapes.pptx")
        .document()
        .expect_err("a PPTX is not a document");
    assert_unexpected(&error, FileFormat::Pptx);
}

#[cfg(all(feature = "docx", feature = "ppt"))]
#[test]
fn a_legacy_ole2_presentation_given_to_the_document_opener_reports_ppt() {
    let error = open("ole/ppt/text_shapes.ppt")
        .document()
        .expect_err("a legacy PowerPoint file is not a document");
    assert_unexpected(&error, FileFormat::Ppt);
}

#[cfg(all(feature = "docx", feature = "odp"))]
#[test]
fn an_odf_presentation_given_to_the_document_opener_reports_odp() {
    let error = open("odf/corpus/impress-basic.odp")
        .document()
        .expect_err("an ODP is not a document");
    assert_unexpected(&error, FileFormat::Odp);
}

#[cfg(feature = "docx")]
#[test]
fn a_non_office_file_given_to_the_document_opener_still_reports_not_an_office_file() {
    assert_not_an_office_file("images/png/lena.png");
}

// ------------------------------------------------------------ presentations --

#[cfg(all(feature = "pptx", feature = "docx"))]
#[test]
fn a_docx_given_to_the_presentation_opener_reports_docx() {
    let error = open("ooxml/docx/comment.docx")
        .presentation()
        .expect_err("a DOCX is not a presentation");
    assert_unexpected(&error, FileFormat::Docx);
}

#[cfg(all(feature = "pptx", feature = "doc"))]
#[test]
fn a_legacy_ole2_document_given_to_the_presentation_opener_reports_doc() {
    let error = open("ole/doc/NoHeadFoot.doc")
        .presentation()
        .expect_err("a legacy Word file is not a presentation");
    assert_unexpected(&error, FileFormat::Doc);
}

#[cfg(all(feature = "pptx", feature = "odt"))]
#[test]
fn an_odf_document_given_to_the_presentation_opener_reports_odt() {
    let error = open("odf/corpus/writer-paragraph-styles.odt")
        .presentation()
        .expect_err("an ODT is not a presentation");
    assert_unexpected(&error, FileFormat::Odt);
}

#[cfg(feature = "pptx")]
#[test]
fn a_non_office_file_given_to_the_presentation_opener_still_reports_not_an_office_file() {
    assert_not_an_office_file("images/png/lena.png");
}

// --------------------------------------------------------------- workbooks --

#[cfg(all(feature = "xlsx", feature = "docx"))]
#[test]
fn a_docx_given_to_the_workbook_opener_reports_docx() {
    let error = open("ooxml/docx/comment.docx")
        .workbook()
        .expect_err("a DOCX is not a workbook");
    assert_unexpected(&error, FileFormat::Docx);
}

#[cfg(all(feature = "xlsx", feature = "pptx"))]
#[test]
fn a_pptx_given_to_the_workbook_opener_reports_pptx() {
    let error = open("ooxml/pptx/shapes.pptx")
        .workbook()
        .expect_err("a PPTX is not a workbook");
    assert_unexpected(&error, FileFormat::Pptx);
}

#[cfg(all(feature = "xlsx", feature = "doc"))]
#[test]
fn a_legacy_ole2_document_given_to_the_workbook_opener_reports_doc() {
    let error = open("ole/doc/NoHeadFoot.doc")
        .workbook()
        .expect_err("a legacy Word file is not a workbook");
    assert_unexpected(&error, FileFormat::Doc);
}

#[cfg(all(feature = "xlsx", feature = "odt"))]
#[test]
fn an_odf_document_given_to_the_workbook_opener_reports_odt() {
    let error = open("odf/corpus/writer-paragraph-styles.odt")
        .workbook()
        .expect_err("an ODT is not a workbook");
    assert_unexpected(&error, FileFormat::Odt);
}

#[cfg(all(feature = "xlsx", feature = "odp"))]
#[test]
fn an_odf_presentation_given_to_the_workbook_opener_reports_odp() {
    let error = open("odf/corpus/impress-basic.odp")
        .workbook()
        .expect_err("an ODP is not a workbook");
    assert_unexpected(&error, FileFormat::Odp);
}

#[cfg(feature = "xlsx")]
#[test]
fn a_non_office_file_given_to_the_workbook_opener_still_reports_not_an_office_file() {
    assert_not_an_office_file("images/png/lena.png");
}

/// The defect this pins: before the change every one of these returned the
/// same unit `NotOfficeFile`, so a caller could not tell a Word document from
/// an unreadable file.
#[cfg(all(
    feature = "xlsx",
    feature = "docx",
    feature = "pptx",
    feature = "doc",
    feature = "odt",
    feature = "odp"
))]
#[test]
fn the_workbook_opener_separates_every_wrong_format_it_can_name() {
    let cases = [
        ("ooxml/docx/comment.docx", FileFormat::Docx),
        ("ooxml/pptx/shapes.pptx", FileFormat::Pptx),
        ("ole/doc/NoHeadFoot.doc", FileFormat::Doc),
        ("odf/corpus/writer-paragraph-styles.odt", FileFormat::Odt),
        ("odf/corpus/impress-basic.odp", FileFormat::Odp),
    ];

    let mut reported = Vec::new();
    for (relative, expected) in cases {
        let error = open(relative)
            .workbook()
            .err()
            .unwrap_or_else(|| panic!("{relative} is not a workbook"));
        assert_unexpected(&error, expected);
        let detected = error.detected().expect("the refusal names the format");
        assert!(
            !reported.contains(&detected),
            "each wrong format must be distinguishable: {detected:?} reported twice"
        );
        reported.push(detected);
    }
    assert_eq!(reported.len(), cases.len());
}
