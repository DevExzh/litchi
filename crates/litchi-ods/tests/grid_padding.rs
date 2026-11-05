use litchi_ods::{Package, Sheet, builder::Workbook as Draft};

const PACKAGE_FIXTURE: &[u8] =
    include_bytes!("../../../test-data/odf/ods/grid-padding-package.ods");

fn signature(sheets: &[Sheet]) -> Vec<(String, usize, usize)> {
    sheets
        .iter()
        .map(|sheet| {
            let populated = sheet
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .filter(|cell| !cell.text.is_empty() || cell.formula.is_some())
                .count();
            (sheet.name.clone(), sheet.rows.len(), populated)
        })
        .collect()
}

#[test]
fn producer_package_opens_without_expanding_grid_padding() {
    let spreadsheet = Package::from_bytes(PACKAGE_FIXTURE)
        .expect("test fixture or operation should succeed")
        .workbook()
        .expect("test fixture or operation should succeed");
    let sheets = spreadsheet.sheets();

    assert_eq!(sheets.len(), 3);
    for sheet in sheets {
        assert!(
            sheet.rows.len() < 4_096,
            "sheet {:?} expanded its padding to {} rows",
            sheet.name,
            sheet.rows.len()
        );
    }
    assert!(
        sheets.iter().any(|sheet| !sheet.rows.is_empty()),
        "the used range must survive"
    );
}

#[test]
fn padded_package_survives_mutable_round_trip() {
    let spreadsheet = Package::from_bytes(PACKAGE_FIXTURE)
        .expect("test fixture or operation should succeed")
        .workbook()
        .expect("test fixture or operation should succeed");
    let before = signature(spreadsheet.sheets());

    let mutable = Draft::from_workbook(spreadsheet);
    let reopened = Package::from_bytes(mutable.build().expect("the draft builds").as_bytes())
        .expect("test fixture or operation should succeed")
        .workbook()
        .expect("test fixture or operation should succeed");

    assert_eq!(signature(reopened.sheets()), before);
}
