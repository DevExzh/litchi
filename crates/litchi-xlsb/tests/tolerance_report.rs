#![allow(
    clippy::unwrap_used,
    reason = "test assertions use panic-on-failure extraction by design"
)]

//! The tolerance report is carried by every open path, and conformant
//! workbooks report clean on each of them.

use std::io::Cursor;

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_xlsb::{Package, SourceBackedWorkbook, ToleranceReport, Workbook};

const CONFORMANT_FIXTURES: &[&str] = &[
    "Simple.xlsb",
    "comments.xlsb",
    "cond_format.xlsb",
    "date.xlsb",
    "hyperlink.xlsb",
    "universal-content.xlsb",
];

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(format!(
        "{}/../../test-data/ooxml/xlsb/{name}",
        env!("CARGO_MANIFEST_DIR")
    ))
    .unwrap()
}

fn assert_clean(report: &ToleranceReport, what: &str) {
    assert!(
        report.is_clean(),
        "{what}: unexpected repairs {:?}",
        report.defects()
    );
    assert_eq!(report.total(), 0, "{what}");
    assert_eq!(report.unrecorded(), 0, "{what}");
    assert_eq!(report.iter().count(), 0, "{what}");
}

#[test]
fn conformant_workbooks_report_clean_on_the_eager_open_path() {
    for name in CONFORMANT_FIXTURES {
        let workbook = Workbook::new(Cursor::new(fixture(name))).unwrap();
        assert_clean(workbook.tolerance_report(), name);
        for index in 0..workbook.worksheet_count() {
            let worksheet = workbook.worksheet(index).unwrap();
            assert_clean(worksheet.tolerance_report(), name);
            let snapshot = workbook.cell_values(index).unwrap();
            assert_clean(snapshot.tolerance_report(), name);
        }
    }
}

#[test]
fn conformant_workbooks_report_clean_on_the_package_open_path() {
    for name in CONFORMANT_FIXTURES {
        let package = Package::from_bytes(fixture(name)).unwrap();
        assert_clean(package.tolerance_report(), name);
        let workbook = package.into_workbook().unwrap();
        assert_clean(workbook.tolerance_report(), name);
    }
}

#[test]
fn conformant_workbooks_report_clean_on_the_source_backed_open_path() {
    for name in CONFORMANT_FIXTURES {
        let workbook = SourceBackedWorkbook::from_reader(Cursor::new(fixture(name))).unwrap();
        assert_clean(&workbook.tolerance_report(), name);
        for worksheet in workbook.worksheets().unwrap() {
            // Some fixtures carry adjunct owners the source-backed layer does
            // not materialize yet; those refusals are outside this contract.
            if let Ok(materialized) = worksheet.materialize() {
                assert_clean(materialized.tolerance_report(), name);
            }
        }
        // Lazily parsed workbook-scope parts (styles) joined the report.
        assert_clean(&workbook.tolerance_report(), name);
    }
}
