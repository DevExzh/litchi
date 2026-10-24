use std::io::Cursor;

use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
use litchi_drawingml::chart::{Limits, chartex, reader, style, writer};
use litchi_drawingml::{Error, LimitKind};

fn chart_xml() -> Vec<u8> {
    Spec::new(Variant::new(Kind::Column))
        .with_categories(["Q1"])
        .with_series(Series::new("Sales", vec![1.0]))
        .to_xml(None)
        .unwrap()
}

fn assert_limit(error: Error, kind: LimitKind) {
    assert!(matches!(error, Error::Limit { kind: actual, .. } if actual == kind));
}

#[test]
fn reader_limits_preflight_nodes_and_depth() {
    let xml = chart_xml();
    let mut limits = Limits::DEFAULT;
    limits.max_xml_nodes = 1;
    assert_limit(
        reader::read(Cursor::new(&xml), &limits).unwrap_err(),
        LimitKind::XmlNodes,
    );

    let mut limits = Limits::DEFAULT;
    limits.max_xml_depth = 1;
    assert_limit(
        reader::read(Cursor::new(&xml), &limits).unwrap_err(),
        LimitKind::XmlDepth,
    );
}

#[test]
fn reader_limits_admit_series_categories_and_points_before_retention() {
    let xml = Spec::new(Variant::new(Kind::Column))
        .with_categories(["Q1", "Q2"])
        .with_series(Series::new("East", vec![1.0, 2.0]))
        .with_series(Series::new("West", vec![3.0, 4.0]))
        .to_xml(None)
        .unwrap();

    let mut limits = Limits::DEFAULT;
    limits.max_series = 1;
    assert_limit(
        reader::read(Cursor::new(&xml), &limits).unwrap_err(),
        LimitKind::ChartSeries,
    );

    let mut limits = Limits::DEFAULT;
    limits.max_categories = 1;
    assert_limit(
        reader::read(Cursor::new(&xml), &limits).unwrap_err(),
        LimitKind::ChartCategories,
    );

    let mut limits = Limits::DEFAULT;
    limits.max_points = 1;
    assert_limit(
        reader::read(Cursor::new(&xml), &limits).unwrap_err(),
        LimitKind::ChartPoints,
    );
}

#[test]
fn authoring_limits_bound_text_before_rebuild() {
    let xml = Spec::new(Variant::new(Kind::Column))
        .with_title("Sales")
        .with_categories(["Q1"])
        .with_series(Series::new("Sales", vec![1.0]))
        .to_xml(None)
        .unwrap();
    let mut limits = Limits::DEFAULT;
    limits.max_text_bytes = 1;
    assert_limit(
        litchi_drawingml::chart::authoring::read(&xml, &limits).unwrap_err(),
        LimitKind::ChartTextBytes,
    );
}

#[test]
fn authoring_limits_bound_rebuilt_output() {
    let xml = chart_xml();
    let parsed = litchi_drawingml::chart::authoring::read(&xml, &Limits::DEFAULT).unwrap();
    let mut limits = Limits::DEFAULT;
    limits.max_output_bytes = 1;
    assert!(!parsed.is_authored(&xml, &limits));
}

#[test]
fn style_limits_bound_strings_and_color_entries() {
    let one_color = br#"<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle"><a:srgbClr val="FF0000"/></cs:colorStyle>"#;
    let mut limits = Limits::DEFAULT;
    limits.max_style_string_bytes = 1;
    assert_limit(
        style::parse_color(one_color, &limits).unwrap_err(),
        LimitKind::ChartStyleStrings,
    );

    let two_colors = br#"<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle"><a:srgbClr val="FF0000"/><a:srgbClr val="00FF00"/></cs:colorStyle>"#;
    let mut limits = Limits::DEFAULT;
    limits.max_style_colors = 1;
    assert_limit(
        style::parse_color(two_colors, &limits).unwrap_err(),
        LimitKind::ChartStyleColors,
    );

    let attributes = br#"<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="1"><a:srgbClr val="FF0000"/></cs:colorStyle>"#;
    let mut limits = Limits::DEFAULT;
    limits.max_style_attributes = 1;
    assert_limit(
        style::parse_color(attributes, &limits).unwrap_err(),
        LimitKind::XmlAttributes,
    );
}

#[test]
fn writer_limits_bound_output_and_reject_invalid_policies() {
    let xml = chart_xml();
    let chart = reader::read(Cursor::new(&xml), &Limits::DEFAULT).unwrap();
    let mut limits = Limits::DEFAULT;
    limits.max_output_bytes = 1;
    assert_limit(
        writer::write(&mut Vec::new(), &chart, &limits).unwrap_err(),
        LimitKind::OutputBytes,
    );

    let mut limits = Limits::DEFAULT;
    limits.max_output_bytes = 0;
    assert!(matches!(
        writer::write(&mut Vec::new(), &chart, &limits),
        Err(Error::InvalidLimit {
            kind: LimitKind::OutputBytes,
            ..
        })
    ));

    let mut invalid = chart;
    invalid.style = Some(49);
    let mut sink = RecordingSink::default();
    assert!(matches!(
        writer::write(&mut sink, &invalid, &Limits::DEFAULT),
        Err(Error::Io(_))
    ));
    assert_eq!(sink.calls, 0);
}

#[test]
fn chartex_limits_bound_points_before_vector_allocation() {
    let xml = Spec::new(Variant::new(Kind::Waterfall))
        .with_categories(["Q1", "Q2"])
        .with_series(Series::new("Cash", vec![1.0, 2.0]))
        .to_xml(None)
        .unwrap();
    let mut limits = Limits::DEFAULT;
    limits.max_points = 1;
    assert_limit(
        chartex::read(&xml, &limits).unwrap_err(),
        LimitKind::ChartPoints,
    );
}

#[test]
fn chartex_limits_bound_series_and_total_points_before_vectors() {
    let xml = Spec::new(Variant::new(Kind::Histogram))
        .with_series(Series::new("East", vec![1.0]))
        .with_series(Series::new("West", vec![2.0]))
        .to_xml(None)
        .unwrap();

    let mut limits = Limits::DEFAULT;
    limits.max_series = 1;
    assert_limit(
        chartex::read(&xml, &limits).unwrap_err(),
        LimitKind::ChartSeries,
    );

    let mut limits = Limits::DEFAULT;
    limits.max_total_points = 1;
    assert_limit(
        chartex::read(&xml, &limits).unwrap_err(),
        LimitKind::ChartPoints,
    );
}

#[derive(Default)]
struct RecordingSink {
    calls: usize,
}

impl std::io::Write for RecordingSink {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.calls += 1;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}
