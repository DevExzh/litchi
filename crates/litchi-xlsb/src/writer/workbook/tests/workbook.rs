#![allow(
    clippy::expect_used,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

//! Workbook-writer serialization and round-trip tests.

use super::super::WorkbookWriter;
use crate::named_ranges::{Definition, area3d_formula};
use crate::writer::MutableWorksheet;
use litchi_core::sheet::{CellValue, WorkbookTrait};
use std::io::Cursor;

#[test]
fn test_create_empty_workbook() {
    let workbook = WorkbookWriter::new();
    assert_eq!(workbook.worksheet_count(), 0);
    assert!(!workbook.is_1904);
}

#[test]
fn test_add_worksheet() {
    let mut workbook = WorkbookWriter::new();
    let sheet = MutableWorksheet::new("Sheet1");
    workbook.add_worksheet(sheet);
    assert_eq!(workbook.worksheet_count(), 1);
}

#[test]
fn test_workbook_writer_default() {
    let workbook: WorkbookWriter = Default::default();
    assert_eq!(workbook.worksheet_count(), 0);
    assert!(!workbook.is_1904);
}

#[test]
fn test_get_worksheet_mut() {
    let mut workbook = WorkbookWriter::new();
    let sheet = MutableWorksheet::new("Sheet1");
    workbook.add_worksheet(sheet);

    let sheet_ref = workbook.get_worksheet_mut(0);
    assert!(sheet_ref.is_some());
    assert_eq!(sheet_ref.unwrap().name(), "Sheet1");

    assert!(workbook.get_worksheet_mut(99).is_none());
}

#[test]
fn test_styles_accessor() {
    let workbook = WorkbookWriter::new();
    let styles = workbook.styles();
    // Just verify it returns a reference
    let _ = styles;
}

#[test]
fn test_styles_mut_accessor() {
    let mut workbook = WorkbookWriter::new();
    let styles = workbook.styles_mut();
    // Just verify it returns a mutable reference
    let _ = styles;
}

#[test]
fn test_add_multiple_worksheets() {
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Sheet1"));
    workbook.add_worksheet(MutableWorksheet::new("Sheet2"));
    workbook.add_worksheet(MutableWorksheet::new("Sheet3"));

    assert_eq!(workbook.worksheet_count(), 3);
}

#[test]
fn test_create_app_xml() {
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Sheet1"));
    workbook.add_worksheet(MutableWorksheet::new("Sheet2"));

    assert_eq!(
        workbook.create_app_xml(),
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>The Litchi Rust Library</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Sheet</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts><Company/><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>"#
    );
}

#[test]
fn test_create_core_xml_is_repeatable_and_does_not_unwind() {
    let workbook = WorkbookWriter::new();
    let first = std::panic::catch_unwind(|| workbook.create_core_xml())
        .expect("deterministic core-property creation must not unwind");
    let second = std::panic::catch_unwind(|| workbook.create_core_xml())
        .expect("repeated core-property creation must not unwind");

    assert_eq!(first, second);
    assert!(!first.contains("<dc:creator"));
    assert!(!first.contains("<cp:lastModifiedBy"));
    assert!(!first.contains("<dcterms:created"));
    assert!(!first.contains("<dcterms:modified"));
    assert!(first.contains("<cp:coreProperties"));
    assert!(first.ends_with("/>"));
}

#[test]
fn test_create_minimal_theme() {
    let workbook = WorkbookWriter::new();
    let theme = workbook.create_minimal_theme();

    assert!(theme.contains("<a:theme"));
    assert!(theme.contains("</a:theme>"));
}

#[test]
fn test_add_named_range() {
    let mut workbook = WorkbookWriter::new();
    let named_range = Definition::new("TestRange".to_string(), None).with_formula(vec![
        crate::package::formula::ptg_types::PTG_INT,
        1,
        0,
    ]);
    workbook.add_named_range(named_range);
    assert_eq!(workbook.named_ranges.len(), 1);
    assert_eq!(workbook.named_ranges[0].name, "TestRange");
}

#[test]
fn defined_name_survives_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Data Sheet"));
    workbook.add_named_range(
        Definition::new("SalesData".to_string(), None)
            .with_formula(area3d_formula(0, 1, 3, 1, 1).unwrap()),
    );
    let mut summary = MutableWorksheet::new("Summary");
    summary.set_cell(
        0,
        0,
        CellValue::Formula {
            formula: "SalesData".to_string(),
            cached_value: Some(Box::new(CellValue::Float(0.0))),
            is_array: false,
            array_range: None,
        },
    );
    summary.set_cell(
        0,
        1,
        CellValue::Formula {
            formula: "'Data Sheet'!$B$2".to_string(),
            cached_value: Some(Box::new(CellValue::Float(0.0))),
            is_array: false,
            array_range: None,
        },
    );
    workbook.add_worksheet(summary);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(reader.defined_names(), &["SalesData"]);
    let summary = reader.worksheet_by_index(1).unwrap();
    assert!(matches!(
        summary.cell_value(0, 0).unwrap().as_ref(),
        CellValue::Formula { formula, .. } if formula == "SalesData"
    ));
    assert!(matches!(
        summary.cell_value(0, 1).unwrap().as_ref(),
        CellValue::Formula { formula, .. } if formula == "'Data Sheet'!$B$2"
    ));
}
