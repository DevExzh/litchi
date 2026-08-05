use litchi_xlsx::Workbook;

#[test]
fn workbook_snapshots_expose_the_first_sheet_by_semantic_name() {
    let workbook = Workbook::create().unwrap();
    assert_eq!(workbook.len(), 1);
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    assert_eq!(sheet.name(), "Sheet1");
}

#[test]
fn workbook_serialization_is_stable_without_pending_edits() {
    let workbook = Workbook::create().unwrap();
    assert_eq!(workbook.to_bytes().unwrap(), workbook.to_bytes().unwrap());
}
