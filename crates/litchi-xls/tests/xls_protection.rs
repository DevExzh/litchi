use std::io::Cursor;

use litchi_xls::raw::Reader;
use litchi_xls::writer::Builder;

#[test]
fn protection_records_round_trip() {
    let mut writer = Builder::new();
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
    writer.build().unwrap().write_to(&mut bytes).unwrap();
    let workbook = Reader::new(Cursor::new(bytes.into_inner())).unwrap();

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
    let mut writer = Builder::new();
    let long_name = "a".repeat(55);
    assert!(
        writer
            .set_file_sharing(false, Some("write"), &long_name)
            .is_err()
    );
}

#[test]
fn worksheet_rename_respects_workbook_structure_but_not_sheet_content_protection() {
    use litchi_xls::{Package, package::edit::Error};
    for structure in [false, true] {
        let mut builder = Builder::new();
        let sheet = builder.add_worksheet("Protected").unwrap();
        builder.protect_sheet(sheet, None, true, true).unwrap();
        if structure {
            builder.protect_workbook(None, true, false);
        }
        let source = builder.build().unwrap();
        let bytes = source.to_bytes().unwrap();
        let p = Package::from_bytes(&bytes).unwrap();
        let mut edit = p.edit().unwrap();
        edit.rename_sheet("Protected", "Protected").unwrap();
        if structure {
            assert!(matches!(edit.rename_sheet("Protected", "Renamed"),
                Err(Error::CellValues(litchi_xls::Error::Refused(r))) if r.kind()==litchi_core::edit::RefusalKind::Protected));
            assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), bytes);
        } else {
            edit.rename_sheet("Protected", "Renamed").unwrap();
            let commit = edit.commit().unwrap();
            assert_eq!(
                commit.package().workbook().worksheet_names().unwrap(),
                ["Renamed"]
            );
            let reopened = Reader::new(Cursor::new(commit.package().to_bytes().unwrap())).unwrap();
            assert!(
                reopened
                    .xls_worksheet(0)
                    .unwrap()
                    .protection()
                    .is_protected()
            );
            assert_eq!(
                commit
                    .package()
                    .apply(&commit.patch().inverse())
                    .unwrap()
                    .package()
                    .to_bytes()
                    .unwrap(),
                bytes
            );
        }
    }
}
