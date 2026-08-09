//! Real-workbook regression coverage for the bounded `BrtCellReal` seam.

#![allow(
    clippy::expect_used,
    reason = "the regression fixture must fail immediately when its checked corpus assumptions break"
)]

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_xlsb::{Workbook, cell_values::Reference};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

#[test]
fn real_fixture_cell_value_commit_is_source_checked_and_length_stable() {
    let (mut workbook, sheet, number) = [
        "test-data/ooxml/xlsb/universal-content.xlsb",
        "test-data/ooxml/xlsb/date.xlsb",
        "test-data/ooxml/xlsb/51519.xlsb",
        "test-data/ooxml/xlsb/62815.xlsb",
        "test-data/ooxml/xlsb/Simple.xlsb",
        "test-data/ooxml/xlsb/cond_format.xlsb",
        "test-data/ooxml/xlsb/comments.xlsb",
        "test-data/ooxml/xlsb/hyperlink.xlsb",
        "test-data/ooxml/xlsb/bug66682.xlsb",
        "test-data/ooxml/xlsb/sample.xlsb",
    ]
    .iter()
    .find_map(|relative| {
        let workbook = Workbook::new(File::open(fixture(relative)).ok()?).ok()?;
        let selected = (0..workbook.worksheet_count()).find_map(|sheet| {
            workbook
                .cell_values(sheet)
                .ok()
                .and_then(|snapshot| snapshot.numbers().find(|number| number.value().is_finite()))
                .map(|number| (sheet, number))
        });
        selected.map(|(sheet, number)| (workbook, sheet, number))
    })
    .expect("a real fixture contains a finite BrtCellReal value");
    let reference =
        Reference::new(number.reference().row(), number.reference().column()).expect("reference");
    let before = workbook.cell_values(sheet).expect("before snapshot");

    let replacement = if number.value().to_bits() == 12_345.25_f64.to_bits() {
        54_321.5
    } else {
        12_345.25
    };
    let mut edit = before.edit();
    edit.set_number(reference, replacement).expect("set number");
    let commit = edit.commit().expect("commit");
    let after = commit.patch().after();
    assert_eq!(before.source_bytes().len(), after.len());
    assert!(
        before
            .source_bytes()
            .iter()
            .zip(after)
            .filter(|(left, right)| left != right)
            .count()
            <= 8,
        "only the selected f64 field may change"
    );

    let published = workbook.apply_cell_values(sheet, &commit).expect("publish");
    assert_eq!(
        published
            .number(reference)
            .expect("lookup")
            .expect("edited value")
            .value()
            .to_bits(),
        replacement.to_bits()
    );
    assert!(commit.patch().apply(before.source_bytes()).is_ok());
    assert!(commit.patch().apply(after).is_err());

    let mut bytes = Cursor::new(Vec::new());
    workbook.save(&mut bytes).expect("save");
    let reopened = Workbook::new(Cursor::new(bytes.into_inner())).expect("reopen");
    assert_eq!(
        reopened
            .cell_values(sheet)
            .expect("reopened snapshot")
            .number(reference)
            .expect("lookup")
            .expect("edited value")
            .value()
            .to_bits(),
        replacement.to_bits()
    );
}
