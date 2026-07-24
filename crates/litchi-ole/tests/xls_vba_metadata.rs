use litchi_ole::xls::{XlsWorkbook, XlsWriter};
use std::fs::File;
use std::io::Cursor;
use std::path::{Path, PathBuf};

fn poi_fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

#[test]
fn reads_macro_fixture_without_opening_project_streams() {
    let workbook = XlsWorkbook::new(File::open(poi_fixture("SimpleMacro.xls")).unwrap()).unwrap();
    let metadata = workbook.vba_metadata();
    assert!(metadata.has_project_marker());
    assert!(metadata.has_project_storage());
    assert!(!metadata.has_no_macros_marker());
    assert!(metadata.may_contain_executable_code());
    assert_eq!(metadata.workbook_code_name(), Some("ThisWorkbook"));
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().vba_code_name(),
        Some("Sheet1")
    );
}

#[test]
fn empty_project_metadata_round_trips_as_non_executable() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Data").unwrap();
    writer.enable_empty_vba_project("ThisWorkbook").unwrap();
    writer
        .set_worksheet_vba_code_name(sheet, Some("DataSheet"))
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let metadata = workbook.vba_metadata();
    assert!(metadata.has_project_marker());
    assert!(metadata.has_no_macros_marker());
    assert!(metadata.has_project_storage());
    assert!(!metadata.may_contain_executable_code());
    assert_eq!(metadata.workbook_code_name(), Some("ThisWorkbook"));
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().vba_code_name(),
        Some("DataSheet")
    );
}

#[test]
fn writer_rejects_invalid_code_names_and_unscoped_sheet_names() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Data").unwrap();
    assert!(
        writer
            .set_worksheet_vba_code_name(sheet, Some("Sheet1"))
            .is_err()
    );
    assert!(writer.enable_empty_vba_project("1Workbook").is_err());
    writer.enable_empty_vba_project("ThisWorkbook").unwrap();
    assert!(
        writer
            .set_worksheet_vba_code_name(sheet, Some("bad-name"))
            .is_err()
    );
}
