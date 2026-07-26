use litchi_ole::xls::{XlsWorkbook, XlsWriter};
use litchi_ole::ovba::{VbaLimits, VbaModuleBuilder, VbaProjectBuilder};
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
    let mut workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
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
    let storage = workbook.vba_project_storage().unwrap();
    assert!(storage.is_structurally_complete());
    assert!(storage.candidate_module_stream_names().is_empty());
    let project = workbook
        .vba_project(&VbaLimits::default())
        .unwrap()
        .unwrap();
    assert_eq!(project.name(), "VBAProject");
    assert!(project.modules().is_empty());
}

#[test]
fn complete_project_with_modules_round_trips_as_inert_source() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Data").unwrap();
    let project = VbaProjectBuilder::new("Analytics")
        .with_module(VbaModuleBuilder::standard(
            "Module1",
            "Public Sub RefreshReport()\r\nEnd Sub\r\n",
        ))
        .with_module(VbaModuleBuilder::document(
            "ThisWorkbook",
            0,
            "Private Sub Workbook_Open()\r\nEnd Sub\r\n",
        ));
    writer
        .set_vba_project("ThisWorkbook", &project, &VbaLimits::default())
        .unwrap();
    writer
        .set_worksheet_vba_code_name(sheet, Some("DataSheet"))
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let mut workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let metadata = workbook.vba_metadata();
    assert!(metadata.has_project_marker());
    assert!(!metadata.has_no_macros_marker());
    assert!(metadata.has_project_storage());
    assert!(metadata.may_contain_executable_code());
    assert_eq!(metadata.workbook_code_name(), Some("ThisWorkbook"));

    let storage = workbook.vba_project_storage().unwrap();
    assert!(storage.is_structurally_complete());
    assert_eq!(
        storage.candidate_module_stream_names(),
        ["Module1", "ThisWorkbook"]
    );
    let project = workbook
        .vba_project(&VbaLimits::default())
        .unwrap()
        .unwrap();
    assert_eq!(project.name(), "Analytics");
    assert_eq!(project.modules().len(), 2);
    assert_eq!(project.modules()[0].name(), "Module1");
    assert!(
        project.modules()[0]
            .source()
            .text()
            .contains("Public Sub RefreshReport()")
    );
    assert_eq!(project.modules()[1].name(), "ThisWorkbook");
    assert!(
        project.modules()[1]
            .source()
            .text()
            .contains("Private Sub Workbook_Open()")
    );
}

#[test]
fn failed_project_build_does_not_replace_existing_configuration() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("Data").unwrap();
    writer.enable_empty_vba_project("ExistingBook").unwrap();
    let project = VbaProjectBuilder::new("TooMany")
        .with_module(VbaModuleBuilder::standard("Module1", "Sub A()\r\nEnd Sub\r\n"));
    let limits = VbaLimits {
        max_modules: 0,
        ..VbaLimits::default()
    };
    assert!(
        writer
            .set_vba_project("ReplacementBook", &project, &limits)
            .is_err()
    );

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    assert_eq!(
        workbook.vba_metadata().workbook_code_name(),
        Some("ExistingBook")
    );
    assert!(workbook.vba_metadata().has_no_macros_marker());
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
