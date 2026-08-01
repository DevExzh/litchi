//! End-to-end coverage for [`XlsOpenOptions::leniency`].
//!
//! Every case opens a real third-party workbook twice: once strictly, proving
//! the defect is still rejected by default, and once leniently, proving the
//! workbook opens *and* that the repair is enumerable rather than silent.

use std::path::PathBuf;

use litchi_xls::{
    XlsFormattingDefect, XlsLeniency, XlsOpenOptions, XlsToleranceReport, XlsWorkbook,
};

fn poi_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

fn libreoffice_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sc/qa/unit/data/xls")
        .join(name)
}

fn lenient_options() -> XlsOpenOptions<'static> {
    XlsOpenOptions {
        password: None,
        leniency: XlsLeniency::TolerateFormattingDefects,
    }
}

fn open_lenient(path: PathBuf) -> XlsWorkbook<std::fs::File> {
    let file = std::fs::File::open(&path).expect("fixture is readable");
    XlsWorkbook::new_with_options(file, lenient_options())
        .expect("a lenient open tolerates cosmetic formatting defects")
}

fn strict_error(path: PathBuf) -> String {
    let file = std::fs::File::open(&path).expect("fixture is readable");
    match XlsWorkbook::new(file) {
        Ok(_) => panic!("a strict open must still reject {}", path.display()),
        Err(error) => error.to_string(),
    }
}

/// Assert that `path` is rejected by default, opens leniently, and reports
/// exactly `defect` as the reason. Returns the first recorded entry.
fn assert_repaired(
    path: PathBuf,
    defect: XlsFormattingDefect,
    strict_needle: &str,
) -> litchi_xls::XlsToleratedDefect {
    let message = strict_error(path.clone());
    assert!(
        message.contains(strict_needle),
        "unexpected strict error for {}: {message}",
        path.display()
    );

    let workbook = open_lenient(path);
    let report = workbook.tolerance_report();
    assert!(!report.is_clean());
    assert_eq!(report.unrecorded(), 0);
    assert_eq!(report.total(), report.defects().len() as u64);
    assert_eq!(
        report.count(defect),
        report.defects().len(),
        "the fixture must isolate a single defect class"
    );
    // The workbook is genuinely usable, not merely constructible.
    assert!(!workbook.sheets().is_empty());

    let entry = report.defects()[0];
    assert_eq!(entry.defect(), defect);
    assert_eq!(entry.record_type(), defect.record_type());
    entry
}

#[test]
fn strict_open_is_the_default_for_every_tolerable_defect() {
    // `XlsWorkbook::new` and a default-constructed options value must agree,
    // and both must reject what a lenient open would repair.
    let path = poi_fixture("29942.xls");
    let explicit = std::fs::File::open(&path).expect("fixture is readable");
    assert!(XlsWorkbook::new_with_options(explicit, XlsOpenOptions::default()).is_err());
    assert!(strict_error(path).contains("Font family"));
}

#[test]
fn lenient_open_repairs_an_out_of_range_font_family() {
    let entry = assert_repaired(
        poi_fixture("29942.xls"),
        XlsFormattingDefect::FontFamily,
        "Font family",
    );
    // The repair substitutes `NotApplicable`; the offending byte survives in
    // the report so a caller can tell exactly what the file claimed.
    assert!(entry.observed() > 5);
}

#[test]
fn lenient_open_repairs_an_empty_font_name() {
    let entry = assert_repaired(
        libreoffice_fixture("tdf170189.xls"),
        XlsFormattingDefect::FontNameEmpty,
        "Font name has 0 characters",
    );
    assert_eq!(entry.observed(), 0);
}

#[test]
fn lenient_open_repairs_an_xfcrc_count_disagreement() {
    let entry = assert_repaired(
        poi_fixture("SharedFormulaTest.xls"),
        XlsFormattingDefect::ExtendedFormatCountMismatch,
        "XFCRC declares",
    );
    // `ordinal` is the number of XF records actually parsed, `observed` the
    // count XFCRC claimed; the two must disagree or there was no defect.
    assert_ne!(entry.ordinal(), entry.observed());
}

#[test]
fn a_conforming_workbook_reports_nothing_under_either_policy() {
    for options in [XlsOpenOptions::default(), lenient_options()] {
        let file = std::fs::File::open(poi_fixture("colwidth.xls")).expect("fixture is readable");
        let workbook =
            XlsWorkbook::new_with_options(file, options).expect("a conforming workbook opens");
        assert!(workbook.tolerance_report().is_clean());
        assert_eq!(workbook.tolerance_report().total(), 0);
    }
}

#[test]
fn a_defect_free_report_is_the_default_state() {
    let report = XlsToleranceReport::default();
    assert!(report.is_clean());
    assert_eq!(report.defects(), &[]);
    assert_eq!(report.count(XlsFormattingDefect::FontNameEmpty), 0);
    assert_eq!(report.total(), 0);
}
