use litchi_xlsx::{Workbook, WorksheetKind};

#[test]
fn workbook_sheet_handles_report_their_semantic_kind() {
    let workbook = Workbook::create().unwrap();
    let sheet = workbook.sheet(0).unwrap().unwrap();
    assert_eq!(sheet.kind(), WorksheetKind::Worksheet);
    assert!(workbook.sheet("Missing").unwrap().is_none());
}
