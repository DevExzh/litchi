use super::codec::MAX_XML_BYTES;
use super::{Mode, Properties, ReferenceMode, parse};
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn parse_xml(child: &str) -> Result<Option<Properties>> {
    parse(format!(r#"<workbook xmlns="{NS}">{child}</workbook>"#).as_bytes())
}

#[test]
fn parses_all_attributes_and_effective_defaults() {
    let value = parse_xml(concat!(
        r#"<calcPr calcId="42" calcMode="autoNoTable" fullCalcOnLoad="1" "#,
        r#"refMode="R1C1" iterate="true" iterateCount="250" iterateDelta="1E-4" "#,
        r#"fullPrecision="0" calcCompleted="false" calcOnSave="0" concurrentCalc="false" "#,
        r#"concurrentManualCount="6" forceFullCalc="true"/>"#,
    ))
    .unwrap()
    .unwrap();
    assert_eq!(value.calculation_id(), 42);
    assert_eq!(value.calculation_mode(), Mode::AutomaticExceptTables);
    assert!(value.full_calculation_on_load());
    assert_eq!(value.reference_mode(), ReferenceMode::R1C1);
    assert!(value.iterative_calculation());
    assert_eq!(value.iteration_count(), 250);
    assert_eq!(value.iteration_delta(), 0.0001);
    assert!(!value.full_precision());
    assert!(!value.calculation_completed());
    assert!(!value.calculate_on_save());
    assert!(!value.concurrent_calculation());
    assert_eq!(value.concurrent_manual_count(), Some(6));
    assert!(value.force_full_calculation());

    let defaults = parse_xml("<calcPr/>").unwrap().unwrap();
    assert_eq!(defaults.calculation_id(), 0);
    assert_eq!(defaults.calculation_mode(), Mode::Automatic);
    assert!(!defaults.full_calculation_on_load());
    assert_eq!(defaults.reference_mode(), ReferenceMode::A1);
    assert!(!defaults.iterative_calculation());
    assert_eq!(defaults.iteration_count(), 100);
    assert_eq!(defaults.iteration_delta(), 0.001);
    assert!(defaults.full_precision());
    assert!(defaults.calculation_completed());
    assert!(defaults.calculate_on_save());
    assert!(defaults.concurrent_calculation());
    assert_eq!(defaults.concurrent_manual_count(), None);
    assert!(!defaults.force_full_calculation());
}

#[test]
fn supports_strict_namespace_and_mce_fallback() {
    let strict = br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><calcPr calcMode="manual"/></workbook>"#;
    assert_eq!(
        parse(strict).unwrap().unwrap().calculation_mode(),
        Mode::Manual
    );
    let mce = format!(
        concat!(
            r#"<workbook xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported">"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x"><x:calcPr/></mc:Choice><mc:Fallback>"#,
            r#"<calcPr refMode="R1C1"/></mc:Fallback></mc:AlternateContent></workbook>"#,
        ),
        NS
    );
    assert_eq!(
        parse(mce.as_bytes()).unwrap().unwrap().reference_mode(),
        ReferenceMode::R1C1
    );
}

#[test]
fn rejects_invalid_values_structure_and_attributes() {
    for child in [
        r#"<calcPr calcMode="sometimes"/>"#,
        r#"<calcPr refMode="A2"/>"#,
        r#"<calcPr iterate="yes"/>"#,
        r#"<calcPr iterateCount="-1"/>"#,
        r#"<calcPr iterateDelta="NaN"/>"#,
        r#"<calcPr iterateDelta="-0.1"/>"#,
        r#"<calcPr mystery="1"/>"#,
        r#"<calcPr><child/></calcPr>"#,
        r#"<wrapper><calcPr calcId="1"/></wrapper>"#,
    ] {
        let result = parse_xml(child);
        if child.starts_with("<wrapper>") {
            assert!(result.unwrap().is_none());
        } else {
            assert!(result.is_err(), "expected rejection for {child}");
        }
    }
    assert!(parse_xml("<calcPr/><calcPr/>").is_err());
    assert!(parse_xml(r#"<calcPr calcId="1" calcId="2"/>"#).is_err());
}

#[test]
fn rejects_entity_references_inside_calc_pr() {
    for reference in ["&amp;", "&#x20;"] {
        let child = format!("<calcPr>{reference}</calcPr>");
        assert!(parse_xml(&child).is_err(), "expected rejection for {child}");
    }
}

#[test]
fn rejects_oversized_workbook_xml() {
    assert!(parse(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
}

fn fixture(bytes: &[u8]) -> Properties {
    let package = OpcPackage::from_bytes(bytes).unwrap();
    let part = package
        .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap();
    parse(part.blob()).unwrap().unwrap()
}

#[test]
fn reads_poi_calculation_fixtures() {
    let iterative = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/47889.xlsx"
    )));
    assert_eq!(iterative.calculation_mode(), Mode::Automatic);
    assert!(iterative.iterative_calculation());
    assert_eq!(iterative.iteration_count(), 100);
    assert_eq!(iterative.iteration_delta(), 0.001);

    let no_save = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/58106.xlsx"
    )));
    assert!(!no_save.calculate_on_save());

    let recalculate = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/60289.xlsx"
    )));
    assert!(recalculate.full_calculation_on_load());
}

#[test]
fn reads_libreoffice_calculation_fixtures() {
    let displayed_precision = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/totalsRowShown.xlsx"
    )));
    assert!(!displayed_precision.full_precision());

    let r1c1 = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf134455.xlsx"
    )));
    assert_eq!(r1c1.reference_mode(), ReferenceMode::R1C1);
    assert_eq!(r1c1.iteration_count(), 100);
    assert_eq!(r1c1.iteration_delta(), 0.001);
}
