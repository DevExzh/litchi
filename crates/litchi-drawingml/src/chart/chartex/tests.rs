use super::*;
use crate::chart::authoring::Grouping;

fn sample(kind: Kind) -> Spec {
    let spec = Spec::new(Variant::new(kind))
        .with_title("Title & <more>")
        .with_categories(["A", "B", "C", "D"]);
    match kind {
        Kind::Histogram | Kind::BoxWhisker => spec
            .with_series(Series::new("First", vec![1.0, 2.0, 2.5, 7.0]))
            .with_series(Series::new("Second", vec![3.0, 4.0, 4.5, 9.0])),
        _ => spec.with_series(Series::new("Only", vec![10.0, -3.0, 4.0, 11.0])),
    }
}

const EXTENDED: [Kind; 7] = [
    Kind::Waterfall,
    Kind::Funnel,
    Kind::Treemap,
    Kind::Sunburst,
    Kind::Histogram,
    Kind::Pareto,
    Kind::BoxWhisker,
];

#[test]
fn every_family_round_trips_through_its_spec() {
    for kind in EXTENDED {
        for binding in [Binding::Literal, Binding::sheet("Sheet1")] {
            let spec = sample(kind).with_binding(binding.clone());
            let external = matches!(binding, Binding::Sheet { .. }).then_some("rId1");
            let xml = spec.to_xml(external).unwrap();
            assert_eq!(xml, spec.to_xml(external).unwrap(), "deterministic");
            let text = String::from_utf8(xml.clone()).unwrap();
            assert!(text.contains(r#"<cx:chartSpace "#), "{kind:?}");
            let chart = read(&xml, &Limits::DEFAULT).unwrap();
            assert_eq!(chart.spec.variant().kind, kind);
            assert_eq!(chart.external_data.as_deref(), external);
            assert_eq!(chart.spec.title(), Some("Title & <more>"));
            assert_eq!(chart.spec.series().len(), spec.series().len());
            assert_eq!(chart.spec.series()[0].values, spec.series()[0].values);
            assert!(
                chart.is_authored(&xml, &Limits::DEFAULT),
                "{kind:?} {binding:?}\n{text}"
            );
            assert_eq!(chart.spec.binding(), &binding, "{kind:?}");
        }
    }
}

#[test]
fn formatting_survives_the_round_trip() {
    let mut appearance = Appearance {
        increase_color: Some("00AA00".into()),
        decrease_color: Some("C00000".into()),
        total_color: Some("804472C4".into()),
        labels: Some(Labels {
            value: true,
            position: Some(DataLabelPosition::OutsideEnd),
            ..Labels::default()
        }),
        axis_min: Some(-10.0),
        axis_max: Some(40.0),
        major_unit: Some(5.0),
        gap_width: Some(80),
        hide_gridlines: true,
        subtotals: Some(vec![0, 3]),
        ..Appearance::default()
    };
    let spec = sample(Kind::Waterfall)
        .with_legend(Some(LegendPosition::TopRight))
        .with_category_axis_title("Quarter")
        .with_value_axis_title("Cash")
        .with_appearance(appearance.clone());
    let xml = spec.to_xml(None).unwrap();
    let chart = read(&xml, &Limits::DEFAULT).unwrap();
    assert!(chart.is_authored(&xml, &Limits::DEFAULT));
    let got = chart.spec.appearance();
    assert_eq!(got.increase_color.as_deref(), Some("00AA00"));
    assert_eq!(got.decrease_color.as_deref(), Some("C00000"));
    assert_eq!(got.total_color.as_deref(), Some("804472C4"));
    assert_eq!(got.subtotals, Some(vec![0, 3]));
    assert_eq!(got.axis_max, Some(40.0));
    assert_eq!(got.gap_width, Some(80));
    assert!(got.hide_gridlines);
    assert_eq!(chart.spec.legend(), Some(LegendPosition::TopRight));
    assert_eq!(chart.spec.category_axis_title(), Some("Quarter"));
    assert_eq!(chart.spec.value_axis_title(), Some("Cash"));

    appearance = Appearance {
        bins: Some(Bins {
            count: Some(4),
            closed_left: true,
            overflow: Some(8.0),
            ..Bins::default()
        }),
        ..Appearance::default()
    };
    let spec = sample(Kind::Histogram).with_appearance(appearance);
    let xml = spec.to_xml(None).unwrap();
    let chart = read(&xml, &Limits::DEFAULT).unwrap();
    assert!(chart.is_authored(&xml, &Limits::DEFAULT));
    assert_eq!(chart.spec.appearance().bins.unwrap().count, Some(4));

    let spec = sample(Kind::Treemap).with_appearance(Appearance {
        colors: vec!["FF0000".into(), "00FF00".into(), "0000FF".into()],
        parent_labels: Some(ParentLabels::Banner),
        ..Appearance::default()
    });
    let xml = spec.to_xml(None).unwrap();
    let chart = read(&xml, &Limits::DEFAULT).unwrap();
    assert!(chart.is_authored(&xml, &Limits::DEFAULT));
    assert_eq!(chart.spec.appearance().colors.len(), 3);

    let spec = sample(Kind::BoxWhisker).with_appearance(Appearance {
        quartile_method: Some(QuartileMethod::Inclusive),
        ..Appearance::default()
    });
    let xml = spec.to_xml(None).unwrap();
    assert!(
        read(&xml, &Limits::DEFAULT)
            .unwrap()
            .is_authored(&xml, &Limits::DEFAULT)
    );
}

#[test]
fn invalid_extended_specs_are_refused() {
    let two = sample(Kind::Waterfall).with_series(Series::new("Extra", vec![1.0]));
    assert!(matches!(two.to_xml(None), Err(Error::Invalid(_))));
    let stacked = Spec::new(Variant::new(Kind::Funnel).with_grouping(Grouping::Stacked))
        .with_series(Series::new("S", vec![1.0]));
    assert!(stacked.to_xml(None).is_err());
    let three_d = Spec::new(Variant::new(Kind::Treemap).with_three_d(true))
        .with_series(Series::new("S", vec![1.0]));
    assert!(three_d.to_xml(None).is_err());
    let both = sample(Kind::Histogram).with_appearance(Appearance {
        bins: Some(Bins {
            count: Some(2),
            size: Some(1.0),
            ..Bins::default()
        }),
        ..Appearance::default()
    });
    assert!(both.to_xml(None).is_err());
    let total = sample(Kind::Waterfall).with_appearance(Appearance {
        subtotals: Some(vec![9]),
        ..Appearance::default()
    });
    assert!(total.to_xml(None).is_err());
    assert!(sample(Kind::Waterfall).to_chart().is_err());
}

#[test]
fn foreign_parts_read_but_are_not_authored() {
    let xml = br#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:foo="urn:foreign"><cx:chartData><cx:data unknown:id="opaque" foo:id="wrong" id="0"><cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt foo:idx="99" idx="0">1</cx:pt><cx:pt foo:idx="99" idx="1">2</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:rich><a:p><a:r><a:t>Rich</a:t></a:r></a:p></cx:rich></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel"><cx:dataId foo:val="wrong" val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#;
    let chart = read(xml, &Limits::DEFAULT).unwrap();
    assert_eq!(chart.spec.variant().kind, Kind::Funnel);
    assert_eq!(chart.spec.title(), Some("Rich"));
    assert_eq!(chart.spec.series()[0].values, [1.0, 2.0]);
    assert!(!chart.is_authored(xml, &Limits::DEFAULT));
    assert!(read(b"<c:chartSpace xmlns:c=\"urn:x\"/>", &Limits::DEFAULT,).is_err());
    assert!(read(
        b"<cx:chartSpace xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chartData/><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId=\"regionMap\"/></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>",
        &Limits::DEFAULT,
    )
    .is_err());
}

#[test]
fn companion_parts_are_valid_chart_styles() {
    let style = crate::chart::style::parse(&style_xml(), &Limits::DEFAULT).unwrap();
    assert_eq!(style.info().entries.len(), 30);
    crate::chart::style::parse_color(&colors_xml(), &Limits::DEFAULT).unwrap();
}

#[test]
fn chartex_writer_checks_external_relationship_limit_before_output() {
    let spec = Spec::new(Variant::new(Kind::Funnel)).with_series(Series::new("Only", vec![1.0]));
    let mut limits = Limits::DEFAULT;
    limits.max_attribute_value_bytes = 2;
    assert!(matches!(
        write_with_limits(&spec, Some("rId3"), &limits),
        Err(Error::Limit {
            kind: crate::LimitKind::AttributeValueBytes,
            ..
        })
    ));
}

#[test]
fn chartex_text_limit_counts_decoded_general_references() {
    let spec = Spec::new(Variant::new(Kind::Funnel))
        .with_title("A&")
        .with_series(Series::new("S", vec![1.0]));
    let xml = spec.to_xml(None).unwrap();

    let mut exact = Limits::DEFAULT;
    exact.max_text_bytes = 2;
    assert!(read(&xml, &exact).is_ok());

    let mut narrow = Limits::DEFAULT;
    narrow.max_text_bytes = 1;
    assert!(matches!(
        read(&xml, &narrow),
        Err(Error::Limit {
            kind: crate::LimitKind::ChartTextBytes,
            ..
        })
    ));
}

#[test]
fn waterfall_subtotals_are_admitted_before_their_vector_grows() {
    let spec = Spec::new(Variant::new(Kind::Waterfall)).with_series(Series::new("Only", vec![1.0]));
    let xml = String::from_utf8(spec.to_xml(None).unwrap())
        .unwrap()
        .replace(
            "</cx:subtotals>",
            r#"<cx:idx val="0"/><cx:idx val="0"/></cx:subtotals>"#,
        );
    let mut limits = Limits::DEFAULT;
    limits.max_points = 2;
    assert!(matches!(
        read(xml.as_bytes(), &limits),
        Err(Error::Limit {
            kind: crate::LimitKind::ChartPoints,
            observed: 3,
            maximum: 2,
        })
    ));
}

#[test]
fn chartex_requires_one_complete_root_and_a_restricted_prolog() {
    let source = sample(Kind::Funnel).to_xml(None).unwrap();
    let mut second_root = source.clone();
    second_root.extend_from_slice(b"<cx:other/>");
    assert!(read(&second_root, &Limits::DEFAULT).is_err());

    let mut text_before_root = b"text".to_vec();
    text_before_root.extend_from_slice(&source);
    assert!(read(&text_before_root, &Limits::DEFAULT).is_err());

    let mut doctype = b"<!DOCTYPE cx:chartSpace>".to_vec();
    doctype.extend_from_slice(&source);
    assert!(read(&doctype, &Limits::DEFAULT).is_err());

    let mut processing_instruction = source;
    processing_instruction.extend_from_slice(b"<?chart extension?>");
    assert!(read(&processing_instruction, &Limits::DEFAULT).is_err());
}
