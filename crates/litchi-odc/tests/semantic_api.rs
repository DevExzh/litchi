#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected errors"
)]

use litchi_odc::{
    AxisSpec, Builder, Chart, ChartClass, ChartClassKind, Definition, Text,
    chart::{Dimension, Kind, Position},
};

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    assert_eq!(Dimension::Z, Dimension::Z);
    assert_eq!(Position::BottomEnd, Position::BottomEnd);

    let bytes = Builder::new().build().unwrap();
    let chart = Chart::from_bytes(bytes).unwrap();
    assert!(!chart.content_xml().contains('\n'));
    assert!(chart.content_xml().contains("<office:chart"));
    assert_eq!(chart.chart().kind(), Kind::Chart);
    assert!(chart.plot_area().is_some());
}

#[test]
fn semantic_whitespace_in_typed_text_is_preserved() {
    let mut definition = Definition::new(ChartClass::line());
    definition.title = Some(Text::new("line one\n  line two"));
    let chart =
        Chart::from_bytes(Builder::new().with_definition(definition).build().unwrap()).unwrap();
    assert!(
        chart
            .content_xml()
            .contains("<text:p>line one\n  line two</text:p>")
    );
}

#[test]
fn validation_is_namespace_aware_and_structural() {
    let mut definition = Definition::new(ChartClass::line());
    definition.plot_area.axes.push(AxisSpec::new(Dimension::X));
    let bytes = Builder::new().with_definition(definition).build().unwrap();
    assert!(
        Chart::from_bytes(bytes)
            .unwrap()
            .chart()
            .plot_area()
            .is_some()
    );

    let mut invalid = Definition::new(ChartClass::line());
    invalid.plot_area.series.push(litchi_odc::SeriesSpec {
        attached_axis: Some("missing".into()),
        ..Default::default()
    });
    assert!(Builder::new().with_definition(invalid).build().is_err());
}

#[test]
fn package_authoring_and_readback_keep_typed_chart_classes_compact() {
    let values = [
        ChartClass::area(),
        ChartClass::bar(),
        ChartClass::bubble(),
        ChartClass::circle(),
        ChartClass::filled_radar(),
        ChartClass::gantt(),
        ChartClass::line(),
        ChartClass::radar(),
        ChartClass::ring(),
        ChartClass::scatter(),
        ChartClass::stock(),
        ChartClass::surface(),
    ];
    for class in values {
        let chart = Chart::from_definition(Definition::new(class.clone())).unwrap();
        assert_eq!(chart.class().unwrap().kind(), class.kind());
        assert!(chart.content_xml().contains(class.lexical()));
        assert!(!chart.content_xml().contains("><\n"));
    }

    let class = ChartClass::extension("vendor:heat", "urn:example:chart").unwrap();
    let chart = Chart::from_definition(Definition::new(class)).unwrap();
    assert_eq!(chart.class().unwrap().kind(), ChartClassKind::Extension);
    assert!(
        chart
            .content_xml()
            .contains("xmlns:vendor=\"urn:example:chart\"")
    );
    assert!(chart.content_xml().contains("chart:class=\"vendor:heat\""));
}

#[test]
fn reader_keeps_chart_class_namespace_alias_and_unknown_value() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><o:body><o:chart><c:chart c:class="c:future"><c:plot-area/></c:chart></o:chart></o:body></o:document-content>"#;
    let chart = litchi_odf_common::chart::read(content).unwrap();
    let class = chart.chart_class().unwrap();
    assert_eq!(class.kind(), ChartClassKind::Unknown);
    assert_eq!(class.lexical(), "c:future");
    assert_eq!(
        class.namespace_uri(),
        Some("urn:oasis:names:tc:opendocument:xmlns:chart:1.0")
    );
}
