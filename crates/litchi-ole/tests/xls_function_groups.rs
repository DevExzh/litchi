use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

use litchi_ole::xls::writer::{XlsFunctionGroupOptions, XlsWriter};
use litchi_ole::xls::{XlsBuiltInFunctionCategories, XlsWorkbook};

#[test]
fn custom_function_groups_round_trip_across_record_generations() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("Functions").unwrap();
    let custom_categories = (0..20)
        .map(|index| {
            if index == 19 {
                "分析".to_string()
            } else {
                format!("Category {index}")
            }
        })
        .collect::<Vec<_>>();
    writer
        .set_function_groups(XlsFunctionGroupOptions {
            built_in: XlsBuiltInFunctionCategories::Sixteen,
            custom_categories: custom_categories.clone(),
        })
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let groups = workbook.function_groups().unwrap();
    assert_eq!(groups.built_in(), XlsBuiltInFunctionCategories::Sixteen);
    assert_eq!(groups.custom_categories(), custom_categories);
    assert_eq!(groups.classic_categories().len(), 16);
    assert_eq!(groups.extended_categories().len(), 4);
    assert_eq!(groups.extended_categories().last().unwrap(), "分析");
}

#[test]
fn reads_poi_and_libreoffice_builtin_function_groups() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let poi = root.join("test-data/poi/test-data/spreadsheet/externalFunctionExample.xls");
    let workbook = XlsWorkbook::new(File::open(poi).unwrap()).unwrap();
    assert_eq!(
        workbook.function_groups().unwrap().built_in(),
        XlsBuiltInFunctionCategories::Sixteen,
    );

    let libreoffice =
        root.join("test-data/libreoffice-core/sc/qa/extras/testdocuments/tdf78897.xls");
    let workbook = XlsWorkbook::new(File::open(libreoffice).unwrap()).unwrap();
    assert_eq!(
        workbook.function_groups().unwrap().built_in(),
        XlsBuiltInFunctionCategories::Fourteen,
    );
}

#[test]
fn reads_established_producer_compatibility_count_seventeen() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/xls/FormulaEvalTestData.xls");
    let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
    assert_eq!(
        workbook.function_groups().unwrap().built_in(),
        XlsBuiltInFunctionCategories::SeventeenCompatibility,
    );
}

#[test]
fn writer_rejects_invalid_function_group_resources() {
    let mut writer = XlsWriter::new();
    assert!(
        writer
            .set_function_groups(XlsFunctionGroupOptions {
                built_in: XlsBuiltInFunctionCategories::Fourteen,
                custom_categories: vec!["A".to_string(), "A".to_string()],
            })
            .is_err()
    );
    assert!(
        writer
            .set_function_groups(XlsFunctionGroupOptions {
                built_in: XlsBuiltInFunctionCategories::Fourteen,
                custom_categories: vec!["x".repeat(33)],
            })
            .is_err()
    );
    assert!(
        writer
            .set_function_groups(XlsFunctionGroupOptions {
                built_in: XlsBuiltInFunctionCategories::Sixteen,
                custom_categories: (0..241).map(|index| index.to_string()).collect(),
            })
            .is_err()
    );
}
