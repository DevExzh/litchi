use std::fs::File;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use litchi_ole::xls::writer::{
    XlsAddInFunctionOptions, XlsDdeOrOleItemOptions, XlsDdeOrOleLinkOptions,
    XlsExternalDefinedNameOptions, XlsExternalSheetOptions, XlsExternalWorkbookOptions, XlsWriter,
};
use litchi_ole::xls::{
    XlsDdeOleValueMatrix, XlsExternalCachedError, XlsExternalCachedValue,
    XlsExternalClipboardFormat, XlsExternalNameBody, XlsSupportingBook, XlsWorkbook,
};

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

#[test]
fn generated_external_names_round_trip_as_inert_metadata() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("Local").unwrap();
    let book = writer
        .add_external_workbook_link(XlsExternalWorkbookOptions {
            encoded_virtual_path: "\u{1}\u{2}Book.xls".to_string(),
            sheets: vec![XlsExternalSheetOptions {
                name: "Remote".to_string(),
                cache_rows: vec![],
            }],
        })
        .unwrap();
    writer
        .add_external_defined_name(
            book,
            XlsExternalDefinedNameOptions {
                name: "RemoteName".to_string(),
                sheet_index: Some(0),
                built_in: false,
                formula_bytes: vec![0x1c, 0x17],
            },
        )
        .unwrap();
    writer
        .add_add_in_function(XlsAddInFunctionOptions {
            name: "ISODD".to_string(),
            unused_data: vec![0x1c, 0x17],
        })
        .unwrap();
    writer
        .add_dde_or_ole_link(XlsDdeOrOleLinkOptions {
            encoded_virtual_path: "\u{6587}\u{3}System".to_string(),
            items: vec![
                XlsDdeOrOleItemOptions {
                    name: "R1C1".to_string(),
                    automatic: true,
                    picture: false,
                    standard_document_name: false,
                    ole_link: false,
                    clipboard_format: XlsExternalClipboardFormat::Text,
                    displayed_as_icon: false,
                    storage_id: 0,
                    matrix: Some(XlsDdeOleValueMatrix {
                        last_column: 1,
                        last_row: 0,
                        values: vec![
                            XlsExternalCachedValue::Text("linked".to_string()),
                            XlsExternalCachedValue::Error(XlsExternalCachedError::NotAvailable),
                        ],
                    }),
                },
                XlsDdeOrOleItemOptions {
                    name: "Object".to_string(),
                    automatic: false,
                    picture: true,
                    standard_document_name: false,
                    ole_link: true,
                    clipboard_format: XlsExternalClipboardFormat::EnhancedMetafile,
                    displayed_as_icon: true,
                    storage_id: 1,
                    matrix: None,
                },
                XlsDdeOrOleItemOptions {
                    name: "StdDocumentName".to_string(),
                    automatic: false,
                    picture: false,
                    standard_document_name: true,
                    ole_link: false,
                    clipboard_format: XlsExternalClipboardFormat::Text,
                    displayed_as_icon: false,
                    storage_id: 0,
                    matrix: None,
                },
            ],
        })
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let links = workbook.external_links();
    assert_eq!(links.supporting_books().len(), 3);
    assert!(matches!(
        links.supporting_books()[0],
        XlsSupportingBook::ExternalWorkbook(_)
    ));
    assert!(matches!(
        links.supporting_books()[1],
        XlsSupportingBook::AddIn
    ));
    assert!(matches!(
        links.supporting_books()[2],
        XlsSupportingBook::DdeOrOle { .. }
    ));
    assert_eq!(links.external_names().len(), 5);
    let XlsExternalNameBody::ExternalDefinedName {
        name,
        sheet_index,
        formula_bytes,
    } = links.external_names()[0].body()
    else {
        panic!("expected external defined name")
    };
    assert_eq!(name, "RemoteName");
    assert_eq!(*sheet_index, Some(0));
    assert_eq!(formula_bytes, &[0x1c, 0x17]);
    let XlsExternalNameBody::DdeOrOle { matrix, .. } = links.external_names()[2].body() else {
        panic!("expected DDE item")
    };
    assert_eq!(matrix.as_ref().unwrap().values.len(), 2);
    assert_eq!(links.sheet_references().len(), 3);
}

#[test]
fn reads_poi_external_and_add_in_names() {
    let workbook = XlsWorkbook::new(File::open(fixture("external_name.xls")).unwrap()).unwrap();
    let names = workbook.external_links().external_names();
    let XlsExternalNameBody::ExternalDefinedName {
        name,
        formula_bytes,
        ..
    } = names[0].body()
    else {
        panic!("expected external name")
    };
    assert_eq!(name, "CreateWeeks");
    assert!(formula_bytes.is_empty());

    let workbook =
        XlsWorkbook::new(File::open(fixture("externalFunctionExample.xls")).unwrap()).unwrap();
    let names = workbook.external_links().external_names();
    assert_eq!(names.len(), 3);
    let XlsExternalNameBody::AddInFunction { name, unused_data } = names[0].body() else {
        panic!("expected add-in function")
    };
    assert_eq!(name, "ISODD");
    assert_eq!(unused_data, &[0x1c, 0x17]);
}
