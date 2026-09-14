//! A wrong-format open reports which format was detected.
//!
//! Detection classifies the input before an opener refuses it, so refusing
//! with a bare "not an Office file" throws away something the caller needs.
//! These tests pin the typed classification onto each opener's refusal, and
//! keep the genuinely unrecognized case reporting `NotOfficeFile`.

#![cfg(any(unix, windows))]

use std::path::{Path, PathBuf};

use litchi::common::Error;
use litchi::common::detection::FileFormat;

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

#[track_caller]
fn assert_unexpected(error: &Error, expected: FileFormat) {
    match error {
        Error::UnexpectedFormat { detected } => assert_eq!(*detected, expected),
        other => panic!("expected UnexpectedFormat {{ {expected:?} }}, got {other:?}"),
    }
}

#[track_caller]
fn assert_boxed_unexpected(
    error: &(dyn std::error::Error + Send + Sync + 'static),
    expected: FileFormat,
) {
    let Some(error) = error.downcast_ref::<Error>() else {
        panic!("expected a litchi error, got {error:?}");
    };
    assert_unexpected(error, expected);
}

// ---------------------------------------------------------------- documents --

#[cfg(all(feature = "docx", feature = "xlsx"))]
#[test]
fn an_xlsx_given_to_the_document_opener_reports_xlsx() {
    let error = litchi::Document::open(fixture("ooxml/xlsx/styles.xlsx"))
        .err()
        .expect("an XLSX is not a document");
    assert_unexpected(&error, FileFormat::Xlsx);
}

#[cfg(all(feature = "docx", feature = "pptx"))]
#[test]
fn a_pptx_given_to_the_document_opener_reports_pptx() {
    let error = litchi::Document::open(fixture("ooxml/pptx/shapes.pptx"))
        .err()
        .expect("a PPTX is not a document");
    assert_unexpected(&error, FileFormat::Pptx);
}

#[cfg(all(feature = "docx", feature = "ppt"))]
#[test]
fn a_legacy_ole2_presentation_given_to_the_document_opener_reports_ppt() {
    let error = litchi::Document::open(fixture("ole/ppt/text_shapes.ppt"))
        .err()
        .expect("a legacy PowerPoint file is not a document");
    assert_unexpected(&error, FileFormat::Ppt);
}

#[cfg(all(feature = "docx", feature = "odp"))]
#[test]
fn an_odf_presentation_given_to_the_document_opener_reports_odp() {
    let error = litchi::Document::open(fixture("odf/corpus/impress-basic.odp"))
        .err()
        .expect("an ODP is not a document");
    assert_unexpected(&error, FileFormat::Odp);
}

#[cfg(feature = "docx")]
#[test]
fn a_non_office_file_given_to_the_document_opener_still_reports_not_an_office_file() {
    let error = litchi::Document::open(fixture("images/png/lena.png"))
        .err()
        .expect("a PNG is not a document");
    assert!(matches!(error, Error::NotOfficeFile), "{error:?}");
}

// ------------------------------------------------------------ presentations --

#[cfg(all(feature = "pptx", feature = "docx"))]
#[test]
fn a_docx_given_to_the_presentation_opener_reports_docx() {
    let error = litchi::Presentation::open(fixture("ooxml/docx/comment.docx"))
        .err()
        .expect("a DOCX is not a presentation");
    assert_unexpected(&error, FileFormat::Docx);
}

#[cfg(all(feature = "pptx", feature = "doc"))]
#[test]
fn a_legacy_ole2_document_given_to_the_presentation_opener_reports_doc() {
    let error = litchi::Presentation::open(fixture("ole/doc/NoHeadFoot.doc"))
        .err()
        .expect("a legacy Word file is not a presentation");
    assert_unexpected(&error, FileFormat::Doc);
}

#[cfg(all(feature = "pptx", feature = "odt"))]
#[test]
fn an_odf_document_given_to_the_presentation_opener_reports_odt() {
    let error = litchi::Presentation::open(fixture("odf/corpus/writer-paragraph-styles.odt"))
        .err()
        .expect("an ODT is not a presentation");
    assert_unexpected(&error, FileFormat::Odt);
}

#[cfg(feature = "pptx")]
#[test]
fn a_non_office_file_given_to_the_presentation_opener_still_reports_not_an_office_file() {
    let error = litchi::Presentation::open(fixture("images/png/lena.png"))
        .err()
        .expect("a PNG is not a presentation");
    assert!(matches!(error, Error::NotOfficeFile), "{error:?}");
}

// --------------------------------------------------------------- workbooks --

#[cfg(all(feature = "xlsx", feature = "docx"))]
#[test]
fn a_docx_given_to_the_workbook_opener_reports_docx() {
    let error = litchi::sheet::Workbook::open(fixture("ooxml/docx/comment.docx"))
        .err()
        .expect("a DOCX is not a workbook");
    assert_boxed_unexpected(error.as_ref(), FileFormat::Docx);
}

#[cfg(all(feature = "xlsx", feature = "pptx"))]
#[test]
fn a_pptx_given_to_the_workbook_opener_reports_pptx() {
    let error = litchi::sheet::Workbook::open(fixture("ooxml/pptx/shapes.pptx"))
        .err()
        .expect("a PPTX is not a workbook");
    assert_boxed_unexpected(error.as_ref(), FileFormat::Pptx);
}

#[cfg(all(feature = "xlsx", feature = "doc"))]
#[test]
fn a_legacy_ole2_document_given_to_the_workbook_opener_reports_doc() {
    let error = litchi::sheet::Workbook::open(fixture("ole/doc/NoHeadFoot.doc"))
        .err()
        .expect("a legacy Word file is not a workbook");
    assert_boxed_unexpected(error.as_ref(), FileFormat::Doc);
}

#[cfg(all(feature = "xlsx", feature = "odt"))]
#[test]
fn an_odf_document_given_to_the_workbook_opener_reports_odt() {
    let error = litchi::sheet::Workbook::open(fixture("odf/corpus/writer-paragraph-styles.odt"))
        .err()
        .expect("an ODT is not a workbook");
    assert_boxed_unexpected(error.as_ref(), FileFormat::Odt);
}

#[cfg(all(feature = "xlsx", feature = "odp"))]
#[test]
fn an_odf_presentation_given_to_the_workbook_opener_reports_odp() {
    let error = litchi::sheet::Workbook::open(fixture("odf/corpus/impress-basic.odp"))
        .err()
        .expect("an ODP is not a workbook");
    assert_boxed_unexpected(error.as_ref(), FileFormat::Odp);
}

#[cfg(feature = "xlsx")]
#[test]
fn a_non_office_file_given_to_the_workbook_opener_still_reports_not_an_office_file() {
    let error = litchi::sheet::Workbook::open(fixture("images/png/lena.png"))
        .err()
        .expect("a PNG is not a workbook");
    let Some(error) = error.downcast_ref::<Error>() else {
        panic!("expected a litchi error, got {error:?}");
    };
    assert!(matches!(error, Error::NotOfficeFile), "{error:?}");
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
        let error = litchi::sheet::Workbook::open(fixture(relative))
            .err()
            .unwrap_or_else(|| panic!("{relative} is not a workbook"));
        assert_boxed_unexpected(error.as_ref(), expected);
        reported.push(expected);
    }

    let mut unique = reported.clone();
    unique.dedup();
    assert_eq!(
        unique.len(),
        reported.len(),
        "each wrong format must be distinguishable"
    );
}
