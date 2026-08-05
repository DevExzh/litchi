//! Opt-in tolerance for non-structural DOC stylesheet defects.

use litchi_doc::{DocLeniency, DocStylesheetDefect, OpenOptions, Package};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc")
        .join(name)
}

/// MS-DOC 2.9 requires style names to be unique, and by default this reader
/// enforces that. The fixture is a real Word document that violates it.
#[test]
fn duplicate_style_names_are_rejected_by_default() {
    let mut package = Package::open(fixture("duplicate-style-names.doc")).expect("container opens");

    match package.document() {
        Err(error) => assert!(
            error.to_string().contains("unique"),
            "unexpected error: {error}"
        ),
        Ok(_) => panic!("a duplicated style name must be fatal under the default contract"),
    }
}

/// Opting in recovers the document. The name is only a label — every stored
/// reference resolves a style by index — so the text and structure are intact.
#[test]
fn opting_in_recovers_the_document_and_records_the_repair() {
    let mut package = Package::open(fixture("duplicate-style-names.doc")).expect("container opens");

    let document = package
        .document_with_options(OpenOptions {
            leniency: DocLeniency::TolerateStylesheetDefects,
            ..Default::default()
        })
        .expect("a duplicated style name is repairable");

    let text = document.text().expect("text is extractable");
    assert!(
        !text.trim().is_empty(),
        "expected the document body to be readable once the stylesheet is tolerated"
    );

    let stylesheet = document.stylesheet().expect("the fixture has a stylesheet");
    let report = stylesheet.tolerance_report();
    assert!(
        !report.is_clean(),
        "the repair must be recorded, never silent"
    );
    assert!(
        report
            .defects()
            .iter()
            .any(|d| d.defect == DocStylesheetDefect::DuplicateStyleName),
        "expected a duplicate-style-name defect, got {:?}",
        report.defects()
    );
    assert_eq!(
        report.total(),
        report.defects().len() as u64,
        "the fixture is far below the recording bound, so nothing should be unrecorded"
    );
}

/// Leniency is scoped to the stylesheet: a structurally broken document stays
/// a hard error in both modes, so opting in cannot mask real corruption.
#[test]
fn structural_defects_stay_fatal_when_tolerating_stylesheet_defects() {
    // Word 6.0 predates the table stream entirely; no stylesheet tolerance can
    // make it readable.
    let mut package = Package::open(fixture("word6-no-table-stream.doc")).expect("container opens");

    assert!(
        package
            .document_with_options(OpenOptions {
                leniency: DocLeniency::TolerateStylesheetDefects,
                ..Default::default()
            })
            .is_err(),
        "leniency must not extend past the stylesheet"
    );
}

/// A conforming document reports a clean slate under either contract.
#[test]
fn a_conforming_document_records_no_repairs() {
    for leniency in [DocLeniency::Strict, DocLeniency::TolerateStylesheetDefects] {
        let mut package = Package::open(fixture("Lists.doc")).expect("container opens");
        let document = package
            .document_with_options(OpenOptions {
                leniency,
                ..Default::default()
            })
            .expect("a conforming document parses under either contract");

        if let Some(stylesheet) = document.stylesheet() {
            assert!(
                stylesheet.tolerance_report().is_clean(),
                "{leniency:?} reported a repair on a conforming document"
            );
        }
    }
}
