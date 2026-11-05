use super::super::write;
use crate::chart::{Chart, ExternalData, HeaderFooter, PivotSource, PrintSettings};

#[test]
fn chart_style_accepts_the_ecma_defined_range() {
    for style in [1, 48] {
        let mut chart = Chart::new();
        chart.style = Some(style);
        write(&mut Vec::new(), &chart, &crate::chart::Limits::DEFAULT).unwrap();
    }
}

#[test]
fn chart_style_rejects_values_outside_the_ecma_defined_range() {
    for style in [0, 49] {
        let mut chart = Chart::new();
        chart.style = Some(style);
        let error = write(&mut Vec::new(), &chart, &crate::chart::Limits::DEFAULT).unwrap_err();
        assert!(matches!(
            error,
            crate::Error::Io(error) if error.kind() == std::io::ErrorKind::InvalidInput
        ));
    }
}

#[test]
fn document_fields_are_checked_before_the_writer_touches_its_sink() {
    let mut limits = crate::chart::Limits::DEFAULT;
    limits.max_text_bytes = 3;
    let mut chart = Chart::new();
    chart.pivot_source = Some(PivotSource::new("long", 1));
    let mut output = Vec::new();
    assert!(matches!(
        write(&mut output, &chart, &limits),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::ChartTextBytes,
            observed: 4,
            maximum: 3,
        })
    ));
    assert!(output.is_empty());

    let mut chart = Chart::new();
    chart.print_settings = Some(PrintSettings {
        header_footer: Some(HeaderFooter {
            odd_header: Some("long".into()),
            ..HeaderFooter::new()
        }),
        ..PrintSettings::new()
    });
    let mut output = Vec::new();
    assert!(matches!(
        write(&mut output, &chart, &limits),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::ChartTextBytes,
            observed: 4,
            maximum: 3,
        })
    ));
    assert!(output.is_empty());

    let mut limits = crate::chart::Limits::DEFAULT;
    limits.max_attribute_value_bytes = 3;
    let mut chart = Chart::new();
    chart.external_data = Some(ExternalData::pending());
    let mut output = Vec::new();
    assert!(matches!(
        super::super::write_with_rels(&mut output, &chart, &limits, Some("rId7"), None,),
        Err(crate::Error::Limit {
            kind: crate::LimitKind::AttributeValueBytes,
            observed: 4,
            maximum: 3,
        })
    ));
    assert!(output.is_empty());
}
