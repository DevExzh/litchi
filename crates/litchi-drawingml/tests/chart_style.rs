use litchi_drawingml::chart::style::{self, ColorMethod, EntryKind};

const CS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

fn entry(name: &str) -> String {
    format!(
        r#"<cs:{name}><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"/></cs:{name}>"#
    )
}

fn chart_style() -> String {
    let names = [
        "axisTitle",
        "categoryAxis",
        "chartArea",
        "dataLabel",
        "dataLabelCallout",
        "dataPoint",
        "dataPoint3D",
        "dataPointLine",
        "dataPointMarker",
        "dataPointWireframe",
        "dataTable",
        "downBar",
        "dropLine",
        "errorBar",
        "floor",
        "gridlineMajor",
        "gridlineMinor",
        "hiLoLine",
        "leaderLine",
        "legend",
        "plotArea",
        "plotArea3D",
        "seriesAxis",
        "seriesLine",
        "title",
        "trendline",
        "trendlineLabel",
        "upBar",
        "valueAxis",
        "wall",
    ];
    let body = names.iter().map(|name| entry(name)).collect::<String>();
    format!(r#"<cs:chartStyle xmlns:cs="{CS}">{body}</cs:chartStyle>"#)
}

#[test]
fn parses_schema_owned_chart_style_without_package_types() {
    let xml = chart_style();
    let document = style::parse(xml.as_bytes()).unwrap();

    assert_eq!(document.info().entries.len(), 30);
    assert_eq!(document.info().entries[0].kind, EntryKind::AxisTitle);
    assert_eq!(document.to_xml(), xml.as_bytes());
}

#[test]
fn parses_color_style_and_rejects_invalid_color() {
    let xml = format!(
        r#"<cs:colorStyle xmlns:cs="{CS}" xmlns:a="{A}" meth="cycle" id="7"><a:srgbClr val="FF0000"><a:shade val="50000"/></a:srgbClr><cs:variation><a:tint val="20000"/></cs:variation></cs:colorStyle>"#
    );
    let document = style::parse_color(xml.as_bytes()).unwrap();

    assert_eq!(document.info().effective_method, ColorMethod::Cycle);
    assert_eq!(document.info().colors.len(), 1);
    assert_eq!(document.info().variations.len(), 1);
    assert!(style::parse_color(xml.replace("FF0000", "GG0000").as_bytes()).is_err());
}
