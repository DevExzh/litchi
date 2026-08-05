use litchi_xls::writer::Writer;
use litchi_xls::{
    Consolidation, ConsolidationBuiltInName, ConsolidationFile, ConsolidationFunction,
    ConsolidationRange, ConsolidationSource, Workbook,
};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

#[test]
fn reads_poi_and_libreoffice_consolidation_directories() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let poi = Workbook::new(
        File::open(root.join("test-data/poi/test-data/spreadsheet/54016.xls")).unwrap(),
    )
    .unwrap();
    let value = poi.xls_worksheet(0).unwrap().consolidation().unwrap();
    assert_eq!(value.function(), ConsolidationFunction::Sum);
    assert!(value.uses_top_labels());
    let ConsolidationSource::CellRange { range, file } = &value.sources()[0] else {
        panic!()
    };
    assert_eq!((range.first_row(), range.last_row()), (1, 8));
    assert_eq!(file.encoded_path(), "\u{2}Sheet1");

    let lo = Workbook::new(
        File::open(
            root.join("test-data/libreoffice-core/sc/qa/extras/testdocuments/NamesSheetLocal.xls"),
        )
        .unwrap(),
    )
    .unwrap();
    let value = (0..lo.sheets().len())
        .find_map(|index| lo.xls_worksheet(index).ok()?.consolidation())
        .unwrap();
    assert_eq!(value.function(), ConsolidationFunction::Sum);
    assert!(value.sources().is_empty());
}

#[test]
fn consolidation_round_trip_preserves_order_and_inert_paths() {
    let mut value = Consolidation::new(ConsolidationFunction::Average);
    value.set_use_left_labels(true);
    value.set_create_links(true);
    value
        .add_source(ConsolidationSource::CellRange {
            range: ConsolidationRange::new(2, 9, 1, 4).unwrap(),
            file: ConsolidationFile::self_reference("Inputs").unwrap(),
        })
        .unwrap();
    value
        .add_source(ConsolidationSource::BuiltInName {
            name: ConsolidationBuiltInName::PrintArea,
            file: None,
        })
        .unwrap();
    value
        .add_source(ConsolidationSource::DefinedName {
            name: "Sales_Data".into(),
            file: Some(ConsolidationFile::new("\u{1}remote.xls\u{3}SheetA").unwrap()),
        })
        .unwrap();
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Output").unwrap();
    writer.set_consolidation(sheet, value).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let value = workbook.xls_worksheet(0).unwrap().consolidation().unwrap();
    assert_eq!(value.function(), ConsolidationFunction::Average);
    assert!(value.uses_left_labels() && value.creates_links());
    assert!(matches!(
        value.sources()[0],
        ConsolidationSource::CellRange { .. }
    ));
    assert!(matches!(
        value.sources()[1],
        ConsolidationSource::BuiltInName { .. }
    ));
    let ConsolidationSource::DefinedName {
        file: Some(file), ..
    } = &value.sources()[2]
    else {
        panic!()
    };
    assert_eq!(file.encoded_path(), "\u{1}remote.xls\u{3}SheetA");
}

#[test]
fn writer_rejects_invalid_names_paths_ranges_and_resources() {
    assert!(ConsolidationFile::new("plain.xls").is_err());
    assert!(ConsolidationRange::new(4, 3, 0, 1).is_err());
    let mut value = Consolidation::new(ConsolidationFunction::Sum);
    assert!(
        value
            .add_source(ConsolidationSource::DefinedName {
                name: "A1".into(),
                file: None
            })
            .is_err()
    );
    assert!(ConsolidationFile::new(format!("\u{1}{}", "x".repeat(4_096))).is_err());
}
