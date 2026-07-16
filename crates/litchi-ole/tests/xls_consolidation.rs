use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;
use litchi_ole::xls::writer::XlsWriter;
use litchi_ole::xls::{
    XlsConsolidation, XlsConsolidationBuiltInName, XlsConsolidationFile,
    XlsConsolidationFunction, XlsConsolidationRange, XlsConsolidationSource, XlsWorkbook,
};

#[test]
fn reads_poi_and_libreoffice_consolidation_directories() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let poi = XlsWorkbook::new(File::open(root.join("3rdparty/poi/test-data/spreadsheet/54016.xls")).unwrap()).unwrap();
    let value = poi.xls_worksheet(0).unwrap().consolidation().unwrap();
    assert_eq!(value.function(), XlsConsolidationFunction::Sum);
    assert!(value.uses_top_labels());
    let XlsConsolidationSource::CellRange { range, file } = &value.sources()[0] else { panic!() };
    assert_eq!((range.first_row(), range.last_row()), (1, 8));
    assert_eq!(file.encoded_path(), "\u{2}Sheet1");

    let lo = XlsWorkbook::new(File::open(root.join("3rdparty/libreoffice-core/sc/qa/extras/testdocuments/NamesSheetLocal.xls")).unwrap()).unwrap();
    let value = (0..lo.sheets().len())
        .find_map(|index| lo.xls_worksheet(index).ok()?.consolidation())
        .unwrap();
    assert_eq!(value.function(), XlsConsolidationFunction::Sum);
    assert!(value.sources().is_empty());
}

#[test]
fn consolidation_round_trip_preserves_order_and_inert_paths() {
    let mut value = XlsConsolidation::new(XlsConsolidationFunction::Average);
    value.set_use_left_labels(true); value.set_create_links(true);
    value.add_source(XlsConsolidationSource::CellRange {
        range: XlsConsolidationRange::new(2, 9, 1, 4).unwrap(),
        file: XlsConsolidationFile::self_reference("Inputs").unwrap(),
    }).unwrap();
    value.add_source(XlsConsolidationSource::BuiltInName { name: XlsConsolidationBuiltInName::PrintArea, file: None }).unwrap();
    value.add_source(XlsConsolidationSource::DefinedName {
        name: "Sales_Data".into(), file: Some(XlsConsolidationFile::new("\u{1}remote.xls\u{3}SheetA").unwrap()),
    }).unwrap();
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Output").unwrap();
    writer.set_consolidation(sheet, value).unwrap();
    let mut bytes = Cursor::new(Vec::new()); writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let value = workbook.xls_worksheet(0).unwrap().consolidation().unwrap();
    assert_eq!(value.function(), XlsConsolidationFunction::Average);
    assert!(value.uses_left_labels() && value.creates_links());
    assert!(matches!(value.sources()[0], XlsConsolidationSource::CellRange { .. }));
    assert!(matches!(value.sources()[1], XlsConsolidationSource::BuiltInName { .. }));
    let XlsConsolidationSource::DefinedName { file: Some(file), .. } = &value.sources()[2] else { panic!() };
    assert_eq!(file.encoded_path(), "\u{1}remote.xls\u{3}SheetA");
}

#[test]
fn writer_rejects_invalid_names_paths_ranges_and_resources() {
    assert!(XlsConsolidationFile::new("plain.xls").is_err());
    assert!(XlsConsolidationRange::new(4, 3, 0, 1).is_err());
    let mut value = XlsConsolidation::new(XlsConsolidationFunction::Sum);
    assert!(value.add_source(XlsConsolidationSource::DefinedName { name: "A1".into(), file: None }).is_err());
    assert!(XlsConsolidationFile::new(format!("\u{1}{}", "x".repeat(4_096))).is_err());
}
