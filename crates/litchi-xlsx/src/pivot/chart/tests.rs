//! Focused pivot-chart parser and package-graph tests.

use super::{C14_CHART_NAMESPACE, DEFAULT_FORMAT_ID};
use super::{FieldType, OPTIONS_EXTENSION_URI, parse_binding};

const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_C: &str = "http://purl.oclc.org/ooxml/drawingml/chart";

#[test]
fn parses_transitional_source_and_options() {
    let xml = format!(
        r#"<c:chartSpace xmlns:c="{C}" xmlns:c14="{C14_CHART_NAMESPACE}">
            <c:pivotSource>
                <c:name>Pivot!PivotOne</c:name>
                <c:fmtId val="7"/>
                <c:extLst><c:ext uri="{{unknown-source}}"/></c:extLst>
            </c:pivotSource>
            <c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/>
                <c:extLst><c:ext uri="{OPTIONS_EXTENSION_URI}">
                    <c14:pivotOptions>
                        <c14:dropZoneVisible val="0"/>
                        <c14:dropZoneData val="1"/>
                    </c14:pivotOptions>
                </c:ext></c:extLst>
            </c:ser></c:barChart></c:plotArea></c:chart>
        </c:chartSpace>"#
    );
    let binding = parse_binding(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(binding.source.name, "Pivot!PivotOne");
    assert_eq!(binding.source.format_id, 7);
    assert_eq!(binding.source.extension_uris, ["{unknown-source}"]);
    assert_eq!(binding.series.len(), 1);
    let options = binding.series[0].options.as_ref().unwrap();
    assert_eq!(options.drop_zone_visible, Some(false));
    assert_eq!(options.visibility(FieldType::DataFields), Some(true));
    assert_eq!(options.visibility(FieldType::AxisRow), None);
}

#[test]
fn parses_strict_chart_namespace_and_ignores_unknown_extensions() {
    let xml = format!(
        r#"<c:chartSpace xmlns:c="{STRICT_C}" xmlns:c14="{C14_CHART_NAMESPACE}">
            <c:pivotSource><c:name>PivotOne</c:name><c:fmtId val="0"/></c:pivotSource>
            <c:chart><c:plotArea><c:barChart><c:ser><c:idx val="3"/>
                <c:extLst>
                    <c:ext uri="{{unknown-series}}"><x:payload xmlns:x="urn:inert"/></c:ext>
                </c:extLst>
            </c:ser></c:barChart></c:plotArea></c:chart>
        </c:chartSpace>"#
    );
    let binding = parse_binding(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(binding.source.name, "PivotOne");
    assert_eq!(binding.series[0].index, 3);
    assert!(binding.series[0].options.is_none());
}

#[test]
fn field_types_and_default_options_are_contextual() {
    assert_eq!("axisRow".parse::<FieldType>().unwrap(), FieldType::AxisRow);
    assert!("bogus".parse::<FieldType>().is_err());
    let options = super::Options::all_visible();
    assert_eq!(options.drop_zones.len(), 5);
    assert!(
        options
            .drop_zones
            .iter()
            .all(|zone| zone.visible && options.visibility(zone.field_type) == Some(true))
    );
    assert_eq!(DEFAULT_FORMAT_ID, 0);
}

#[test]
fn resolves_authored_source_names_without_leaking_package_details() {
    let tables = vec![("PivotOne".to_string(), "Data".to_string())];
    assert_eq!(
        super::resolve_source_name("PivotOne", "Other", &tables).unwrap(),
        "Data!PivotOne"
    );
    assert!(super::resolve_source_name("Missing", "Other", &tables).is_err());
}

#[test]
fn ordinary_chart_has_no_binding() {
    let xml = format!(
        r#"<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:barChart>
            <c:ser><c:idx val="0"/></c:ser></c:barChart></c:plotArea>
        </c:chart></c:chartSpace>"#
    );
    assert!(parse_binding(xml.as_bytes()).unwrap().is_none());
}
