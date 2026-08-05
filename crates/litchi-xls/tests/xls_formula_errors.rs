use std::fs::File;

use litchi_xls::Workbook;

#[test]
fn parses_poi_formula_error_shared_feature_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/46136-NoWarnings.xls"
    );
    let workbook = Workbook::new(File::open(path).unwrap()).unwrap();
    let worksheet = workbook.xls_worksheet(0).unwrap();
    let features = worksheet.formula_error_features();
    assert_eq!(features.len(), 1);
    assert_eq!(features[0].ranges().len(), 1);
    let range = features[0].ranges()[0];
    assert_eq!((range.first_row(), range.last_row()), (0, 0));
    assert_eq!((range.first_column(), range.last_column()), (0, 0));
    assert_eq!(features[0].checks().bits(), 0x04);
    assert!(features[0].checks().numbers_stored_as_text());
}
