use litchi_xlsx::WorksheetKind;

#[test]
fn workbook_sheet_handles_report_their_semantic_kind() {
    let workbook = litchi_xlsx::Package::new().unwrap();
    let sheet = workbook.workbook_ref().sheet(0).unwrap().unwrap();
    assert_eq!(sheet.kind(), WorksheetKind::Worksheet);
    assert!(workbook.workbook_ref().sheet("Missing").unwrap().is_none());
}
