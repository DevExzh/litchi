use litchi_ods::{CellValue, Spreadsheet};
use std::path::{Path, PathBuf};

fn corpus() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/odf/corpus")
}

#[test]
fn real_calc_packages_parse_and_expose_typed_content() {
    let cases = [
        ("calc-formulas.ods", 3),
        ("calc-unicode-chinese.ods", 1),
        ("calc-cell-styles.ods", 1),
        ("calc-two-sheets.ods", 2),
    ];
    for (file, count) in cases {
        let spreadsheet = Spreadsheet::open(corpus().join(file))
            .unwrap_or_else(|error| panic!("{file}: open failed: {error:?}"));
        assert_eq!(spreadsheet.sheets().len(), count, "{file}: sheet count");
        assert!(
            spreadsheet
                .sheets()
                .iter()
                .any(|sheet| !sheet.rows.is_empty()),
            "{file}: no populated rows"
        );
        assert!(
            spreadsheet
                .sheets()
                .iter()
                .flat_map(|sheet| &sheet.rows)
                .flat_map(|row| &row.cells)
                .any(|cell| !cell.text.is_empty() || cell.value != CellValue::Empty),
            "{file}: no extractable cell content"
        );
        let _ = spreadsheet.metadata();
    }

    let spreadsheet = Spreadsheet::open(corpus().join("calc-formulas.ods")).unwrap();
    assert!(
        spreadsheet
            .sheets()
            .iter()
            .flat_map(|sheet| &sheet.rows)
            .flat_map(|row| &row.cells)
            .any(|cell| cell.formula.is_some()),
        "formula fixture must expose at least one formula cell"
    );
}
