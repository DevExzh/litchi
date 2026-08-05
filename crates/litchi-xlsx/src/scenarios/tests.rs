//! Regression tests for the worksheet scenario owner.

use super::model::{MAX_DEPTH, MAX_XSTRING_CHARS};
use super::*;
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn parse(child: &str) -> Result<Option<Collection>> {
    parse_worksheet_scenarios(format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes())
}

#[test]
fn parses_scenarios_attributes_input_cells_and_defaults() {
    let value = parse(concat!(
        r#"<scenarios current="1" show="1" sqref="A1 $B$2:C3">"#,
        r#"<scenario name="first"><inputCells r="A1" val="10"/></scenario>"#,
        r#"<scenario name="second" locked="1" hidden="true" count="2" user="one" comment="note">"#,
        r#"<inputCells r="B2" val="x" numFmtId="14"/>"#,
        r#"<inputCells r="$C$3" deleted="1" undone="true" val="y &amp; z"/></scenario>"#,
        r#"</scenarios>"#,
    ))
    .unwrap()
    .unwrap();
    assert_eq!(value.current(), Some(1));
    assert_eq!(value.show(), Some(1));
    assert_eq!(
        value
            .ranges()
            .iter()
            .map(RangeReference::as_str)
            .collect::<Vec<_>>(),
        vec!["A1", "$B$2:C3"]
    );
    assert_eq!(value.scenarios().len(), 2);
    let first = &value.scenarios()[0];
    assert_eq!(first.name(), "first");
    assert!(!first.locked());
    assert!(!first.hidden());
    assert_eq!(first.count(), None);
    assert_eq!(first.user(), None);
    assert_eq!(first.input_cells()[0].reference().as_str(), "A1");
    assert_eq!(first.input_cells()[0].value(), "10");
    let second = &value.scenarios()[1];
    assert!(second.locked());
    assert!(second.hidden());
    assert_eq!(second.count(), Some(2));
    assert_eq!(second.user(), Some("one"));
    assert_eq!(second.comment(), Some("note"));
    let cells = second.input_cells();
    assert_eq!(cells[0].number_format_id(), Some(14));
    assert!(cells[1].deleted());
    assert!(cells[1].undone());
    assert_eq!(cells[1].value(), "y & z");
}

#[test]
fn supports_strict_namespace_and_skips_extension_markup() {
    let xml = concat!(
        r#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">"#,
        r#"<scenarios><scenario name="s">"#,
        r#"<extLst><ext uri="urn:test"><x:payload xmlns:x="urn:x"/></ext></extLst>"#,
        r#"</scenario></scenarios></worksheet>"#,
    );
    let value = parse_worksheet_scenarios(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(value.scenarios()[0].name(), "s");
}

#[test]
fn rejects_structure_attributes_and_limits() {
    for child in [
        "<scenarios/>",
        r#"<scenarios><scenario/></scenarios>"#,
        r#"<scenarios><scenario name=""><inputCells val="1"/></scenario></scenarios>"#,
        r#"<scenarios><scenario name="s"><inputCells r="A1"/></scenario></scenarios>"#,
        r#"<scenarios><scenario name="s"><inputCells r="A0" val="1"/></scenario></scenarios>"#,
        r#"<scenarios><scenario name="s"><inputCells r="XFE1" val="1"/></scenario></scenarios>"#,
        r#"<scenarios current="yes"><scenario name="s"/></scenarios>"#,
        r#"<scenarios sqref=""><scenario name="s"/></scenarios>"#,
        r#"<scenarios><scenario name="s" locked="yes"/></scenarios>"#,
        r#"<scenarios><scenario name="s">text</scenario></scenarios>"#,
    ] {
        assert!(parse(child).is_err(), "expected rejection for {child}");
    }
    assert!(parse("<scenarios><scenario name=\"s\"/></scenarios><scenarios><scenario name=\"t\"/></scenarios>").is_err());
    let long_name = "x".repeat(MAX_XSTRING_CHARS + 1);
    assert!(
        parse(&format!(
            r#"<scenarios><scenario name="{long_name}"/></scenarios>"#
        ))
        .is_err()
    );
}

#[test]
fn preserves_unknown_attributes_and_elements() {
    let xml = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="urn:scenario-extension">"#,
        r#"<scenarios x:collection="one"><x:before/><scenario name="s" x:scenario="two">"#,
        r#"<inputCells r="A1" val="1" x:cell="three"><x:inside/></inputCells>"#,
        r#"<x:after/></scenario><x:tail/></scenarios></worksheet>"#,
    );
    let value = parse_worksheet_scenarios(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(value.unknown_attributes()[0].name(), "x:collection");
    assert_eq!(value.unknown_attributes()[0].value(), "one");
    assert_eq!(value.unknown_elements().len(), 2);
    let scenario = &value.scenarios()[0];
    assert_eq!(scenario.unknown_attributes()[0].name(), "x:scenario");
    assert_eq!(scenario.unknown_elements().len(), 1);
    let cell = &scenario.input_cells()[0];
    assert_eq!(cell.unknown_attributes()[0].name(), "x:cell");
    assert_eq!(cell.unknown_elements().len(), 1);

    let fragment = write_worksheet_scenarios(&value, Conformance::Transitional).unwrap();
    let document = format!(r#"<worksheet xmlns="{NS}">{fragment}</worksheet>"#);
    let reparsed = parse_worksheet_scenarios(document.as_bytes())
        .unwrap()
        .unwrap();
    assert_eq!(reparsed, value);
}

#[test]
fn rejects_malformed_document_boundaries_and_excessive_depth() {
    for xml in [
        format!(r#"<worksheet xmlns="{NS}"/><worksheet xmlns="{NS}"/>"#),
        format!(r#"text<worksheet xmlns="{NS}"></worksheet>"#),
        format!(r#"<worksheet xmlns="{NS}">text</worksheet>"#),
        format!(r#"<worksheet xmlns="{NS}"></worksheet>tail"#),
        format!(r#"<worksheet xmlns="{NS}"><![CDATA[data]]></worksheet>"#),
        format!(
            r#"<worksheet xmlns="{NS}"><scenarios><scenario name="s"/></scenarios></worksheet><?pi?>"#
        ),
    ] {
        assert!(
            parse_worksheet_scenarios(xml.as_bytes()).is_err(),
            "expected rejection for {xml}"
        );
    }

    let mut xml = format!(r#"<worksheet xmlns="{NS}">"#);
    for _ in 0..MAX_DEPTH {
        xml.push_str("<extension>");
    }
    for _ in 0..MAX_DEPTH {
        xml.push_str("</extension>");
    }
    xml.push_str("</worksheet>");
    assert!(parse_worksheet_scenarios(xml.as_bytes()).is_err());
}

#[test]
fn write_round_trips_through_the_reader() {
    let scenario = Scenario::new("baseline")
        .unwrap()
        .with_locked(true)
        .with_count(1)
        .with_user("analyst")
        .unwrap()
        .with_comment("Q1 <plan> & \"notes\"")
        .unwrap()
        .with_input_cells(vec![
            InputCell::new(CellReference::new("A1").unwrap(), "10")
                .unwrap()
                .with_number_format_id(14),
            InputCell::new(CellReference::new("$B$2").unwrap(), "hold")
                .unwrap()
                .with_deleted(true)
                .with_undone(true),
        ])
        .unwrap();
    let expected = Collection::new(vec![scenario])
        .unwrap()
        .with_current(0)
        .with_show(0)
        .with_ranges(vec![RangeReference::new("A1:B2").unwrap()])
        .unwrap();
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let fragment = write_worksheet_scenarios(&expected, conformance).unwrap();
        let document = format!(r#"<worksheet xmlns="{NS}">{fragment}</worksheet>"#);
        let parsed = parse_worksheet_scenarios(document.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(parsed, expected);
    }
}

#[test]
fn writer_rejects_output_over_limit_before_final_append() {
    let long_name = "x".repeat(MAX_XSTRING_CHARS);
    let scenarios = (0..1_025)
        .map(|_| Scenario::new(long_name.clone()).unwrap())
        .collect();
    let value = Collection::new(scenarios).unwrap();
    let error = write_worksheet_scenarios(&value, Conformance::Transitional).unwrap_err();
    assert!(
        error
            .to_string()
            .contains("serialized scenarios XML exceeds safety limit")
    );
}

#[test]
fn reads_libreoffice_scenario_fixture() {
    let package = OpcPackage::from_bytes(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf97598_scenarios.xlsx"
    )))
    .unwrap();
    let part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap();
    let value = parse_worksheet_scenarios(part.blob()).unwrap().unwrap();
    assert_eq!(value.current(), Some(0));
    assert_eq!(value.scenarios().len(), 1);
    let scenario = &value.scenarios()[0];
    assert_eq!(scenario.name(), "scenario1");
    assert!(scenario.locked());
    assert_eq!(scenario.count(), Some(1));
    assert_eq!(scenario.user(), Some("one"));
    assert_eq!(scenario.comment(), Some("Created by one on 12/26/2016"));
    assert_eq!(scenario.input_cells().len(), 1);
    assert_eq!(scenario.input_cells()[0].reference().as_str(), "A1");
    assert_eq!(scenario.input_cells()[0].value(), "value from scenario1");
    // Sheets without scenarios parse to None.
    let other = package
        .get_part(&PackURI::new("/xl/worksheets/sheet2.xml").unwrap())
        .unwrap();
    assert!(parse_worksheet_scenarios(other.blob()).unwrap().is_none());
}
