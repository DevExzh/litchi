use litchi_ole::xls::writer::XlsWorkbookEnvironmentOptions;
use litchi_ole::xls::{XlsLinkUpdateMode, XlsObjectDisplayMode, XlsWorkbook, XlsWriter};
use std::fs::File;
use std::io::Cursor;
use std::path::{Path, PathBuf};

fn libreoffice_fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sc/qa/unit/data/xls")
        .join(name)
}

fn poi_fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

#[test]
fn reads_libreoffice_environment_fixture_without_refreshing_data() {
    let workbook =
        XlsWorkbook::new(File::open(libreoffice_fixture("formats.xls")).unwrap()).unwrap();
    let environment = workbook.environment();
    assert!(environment.has_excel9_file_marker());
    assert!(environment.supports_natural_language_formulas());
    assert!(!environment.refresh_external_data_on_load());
    assert_eq!(
        environment.object_display_mode(),
        XlsObjectDisplayMode::ShowAll
    );
    assert_eq!(environment.default_country_code(), 1);
}

#[test]
fn workbook_environment_round_trip() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("Environment").unwrap();
    writer
        .set_workbook_environment(XlsWorkbookEnvironmentOptions {
            template: true,
            has_biff5_stream: true,
            create_backup_copy: true,
            object_display_mode: XlsObjectDisplayMode::HideAll,
            refresh_external_data_on_load: true,
            save_external_link_values: false,
            has_envelope: true,
            envelope_visible: true,
            envelope_initialized: true,
            link_update_mode: XlsLinkUpdateMode::Silent,
            hide_unselected_table_borders: true,
            supports_natural_language_formulas: true,
            default_country_code: 44,
            current_country_code: 86,
        })
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let environment = workbook.environment();
    assert!(environment.is_template());
    assert!(environment.has_biff5_stream());
    assert!(environment.create_backup_copy());
    assert_eq!(
        environment.object_display_mode(),
        XlsObjectDisplayMode::HideAll
    );
    assert!(environment.refresh_external_data_on_load());
    assert!(!environment.save_external_link_values());
    assert_eq!(environment.link_update_mode(), XlsLinkUpdateMode::Silent);
    assert_eq!(environment.current_country_code(), 86);
}

#[test]
fn reads_poi_dual_stream_dsf_value() {
    let workbook =
        XlsWorkbook::new(File::open(poi_fixture("SimpleWithPrintArea.xls")).unwrap()).unwrap();
    assert!(workbook.environment().has_biff5_stream());
}

#[test]
fn writer_rejects_cross_field_and_country_bounds() {
    let mut writer = XlsWriter::new();
    let refresh = XlsWorkbookEnvironmentOptions {
        refresh_external_data_on_load: true,
        ..XlsWorkbookEnvironmentOptions::default()
    };
    assert!(writer.set_workbook_environment(refresh).is_err());
    let envelope = XlsWorkbookEnvironmentOptions {
        envelope_visible: true,
        ..XlsWorkbookEnvironmentOptions::default()
    };
    assert!(writer.set_workbook_environment(envelope).is_err());
    let country = XlsWorkbookEnvironmentOptions {
        default_country_code: 0,
        ..XlsWorkbookEnvironmentOptions::default()
    };
    assert!(writer.set_workbook_environment(country).is_err());
}
