use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

use litchi_ole::xls::writer::{
    XlsExternalCacheRowOptions, XlsExternalSheetOptions, XlsExternalWorkbookOptions, XlsWriter,
};
use litchi_ole::xls::{XlsExternalCachedError, XlsExternalCachedValue, XlsWorkbook};

#[test]
fn external_workbook_cache_round_trip_is_inert() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("Local").unwrap();
    writer.add_external_workbook_link(XlsExternalWorkbookOptions {
        encoded_virtual_path: "\u{1}remote-data.xls".to_string(),
        sheets: vec![XlsExternalSheetOptions {
            name: "Inputs".to_string(),
            cache_rows: vec![XlsExternalCacheRowOptions {
                row: 7,
                first_column: 2,
                values: vec![
                    XlsExternalCachedValue::Blank,
                    XlsExternalCachedValue::Number(42.5),
                    XlsExternalCachedValue::Text("cached".to_string()),
                    XlsExternalCachedValue::Boolean(true),
                    XlsExternalCachedValue::Error(XlsExternalCachedError::NotAvailable),
                ],
            }],
        }],
    }).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let external = workbook.external_links().external_workbooks().next().unwrap();
    assert_eq!(external.encoded_virtual_path(), "\u{1}remote-data.xls");
    assert_eq!(external.sheets()[0].name(), "Inputs");
    assert_eq!(external.sheets()[0].cache_rows()[0].values().len(), 5);
    assert_eq!(
        external.sheets()[0].cache_rows()[0].values()[2],
        XlsExternalCachedValue::Text("cached".to_string()),
    );
}

#[test]
fn reads_poi_and_libreoffice_external_caches() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let poi = root.join("3rdparty/poi/test-data/spreadsheet/XRefCalc.xls");
    let workbook = XlsWorkbook::new(File::open(poi).unwrap()).unwrap();
    let external = workbook.external_links().external_workbooks().next().unwrap();
    assert_eq!(external.sheets().len(), 3);
    assert_eq!(external.sheets()[0].name(), "MarkupSheet");
    assert!(!external.sheets()[0].cache_rows().is_empty());

    let libreoffice = root.join("3rdparty/libreoffice-core/sc/qa/unit/data/xls/external-ref.xls");
    let workbook = XlsWorkbook::new(File::open(libreoffice).unwrap()).unwrap();
    let external = workbook.external_links().external_workbooks().next().unwrap();
    assert_eq!(external.sheets()[0].name(), "Sheet1");
    assert_eq!(
        external.sheets()[0].cache_rows()[0].values(),
        &[XlsExternalCachedValue::Text("external text".to_string())],
    );
}

#[test]
fn writer_rejects_external_cache_resource_errors() {
    let mut writer = XlsWriter::new();
    assert!(writer.add_external_workbook_link(XlsExternalWorkbookOptions {
        encoded_virtual_path: String::new(),
        sheets: vec![],
    }).is_err());
    assert!(writer.add_external_workbook_link(XlsExternalWorkbookOptions {
        encoded_virtual_path: "\u{1}book.xls".to_string(),
        sheets: vec![XlsExternalSheetOptions {
            name: "Bad/Name".to_string(),
            cache_rows: vec![],
        }],
    }).is_err());
}
