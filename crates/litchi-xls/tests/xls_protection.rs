use std::io::Cursor;

use litchi_xls::XlsWorkbook;
use litchi_xls::writer::XlsWriter;

#[test]
fn protection_records_round_trip() {
    let mut writer = XlsWriter::new();
    let sheet_index = writer.add_worksheet("Protected").unwrap();
    writer.protect_workbook(Some("book"), true, false);
    writer.protect_revisions(Some("revision"));
    writer
        .protect_sheet(sheet_index, Some("sheet"), true, true)
        .unwrap();
    writer
        .set_file_sharing(true, Some("write"), "审阅者")
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();

    let protection = workbook.protection();
    assert!(protection.structure_protected());
    assert!(!protection.windows_protected());
    assert!(protection.password().is_set());
    assert!(protection.revisions_protected());
    assert!(protection.revision_password().is_set());
    assert!(protection.write_protected());
    let sharing = protection.file_sharing().unwrap();
    assert!(sharing.read_only_recommended());
    assert!(sharing.write_password().is_set());
    assert_eq!(sharing.user_name(), "审阅者");

    let sheet = workbook.xls_worksheet(0).unwrap();
    assert!(sheet.protection().is_protected());
    assert!(sheet.protection().objects_protected());
    assert!(sheet.protection().scenarios_protected());
    assert!(sheet.protection().has_password());
}

#[test]
fn file_sharing_rejects_long_user_name() {
    let mut writer = XlsWriter::new();
    let long_name = "a".repeat(55);
    assert!(
        writer
            .set_file_sharing(false, Some("write"), &long_name)
            .is_err()
    );
}
