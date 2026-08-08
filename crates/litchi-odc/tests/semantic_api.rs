#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected errors"
)]

use litchi_odc::{
    AxisSpec, Builder, Chart, Definition, Text,
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
    let mut definition = Definition::new("chart:line");
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
    let mut definition = Definition::new("chart:line");
    definition.plot_area.axes.push(AxisSpec::new(Dimension::X));
    let bytes = Builder::new().with_definition(definition).build().unwrap();
    assert!(
        Chart::from_bytes(bytes)
            .unwrap()
            .chart()
            .plot_area()
            .is_some()
    );

    let mut invalid = Definition::new("chart:line");
    invalid.plot_area.series.push(litchi_odc::SeriesSpec {
        attached_axis: Some("missing".into()),
        ..Default::default()
    });
    assert!(Builder::new().with_definition(invalid).build().is_err());
}
