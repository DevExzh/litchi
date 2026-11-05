use super::*;
use crate::chart::edit::Package;
use crate::chart::reader;

fn sales(variant: Variant) -> Spec {
    Spec::new(variant)
        .with_title("Sales")
        .with_categories(["Q1", "Q2", "Q3", "Q4"])
        .with_series(Series::new("East", vec![120.0, 135.0, 148.0, 162.0]))
        .with_series(Series::new("West", vec![95.0, 108.0, 115.0, 128.0]))
}

fn text(title: Option<&TitleText>) -> Option<String> {
    match title? {
        TitleText::Literal(rich) => Some(rich.text.clone()),
        TitleText::Reference(_) => None,
    }
}

#[test]
fn every_variant_round_trips_through_the_reader() {
    let kinds = [
        Kind::Column,
        Kind::Bar,
        Kind::Line,
        Kind::Pie,
        Kind::Doughnut,
        Kind::Area,
        Kind::Scatter,
        Kind::Bubble,
        Kind::Radar,
    ];
    for kind in kinds {
        for grouping in [
            Grouping::Standard,
            Grouping::Stacked,
            Grouping::PercentStacked,
        ] {
            for three_d in [false, true] {
                let variant = Variant::new(kind)
                    .with_grouping(grouping)
                    .with_three_d(three_d);
                let mut spec = sales(variant);
                if matches!(kind, Kind::Pie | Kind::Doughnut) {
                    spec.series.truncate(1);
                }
                if !variant.is_valid() {
                    assert!(spec.to_xml(None).is_err(), "{variant:?}");
                    continue;
                }
                for binding in [Binding::Literal, Binding::sheet("Sheet1")] {
                    let spec = spec.clone().with_binding(binding.clone());
                    let xml = spec.to_xml(Some("rId1")).unwrap();
                    let chart = reader::read(&xml[..], &crate::chart::Limits::DEFAULT).unwrap();
                    assert_eq!(
                        Variant::of(&chart),
                        Some(variant),
                        "{variant:?} {binding:?}"
                    );
                    assert_eq!(text(chart.title.as_ref()).as_deref(), Some("Sales"));
                    let series = &chart.plot_area.type_groups[0].common().series;
                    assert_eq!(series.len(), spec.series().len());
                    // Deterministic: the same spec writes the same bytes.
                    assert_eq!(xml, spec.to_xml(Some("rId1")).unwrap());
                    // The source-preserving editor accepts every authored part.
                    Package::from_xml(xml).unwrap();
                }
            }
        }
    }
}

#[test]
fn stock_charts_need_three_or_four_series() {
    let spec = sales(Variant::new(Kind::Stock));
    assert!(spec.validate().is_err());
    let spec = spec.with_series(Series::new("Close", vec![1.0, 2.0, 3.0, 4.0]));
    let chart = reader::read(
        &spec.to_xml(None).unwrap()[..],
        &crate::chart::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(Variant::of(&chart), Some(Variant::new(Kind::Stock)));
}

#[test]
fn sheet_binding_refers_to_the_documented_layout() {
    let spec = sales(Variant::new(Kind::Column)).with_binding(Binding::sheet("Sheet1"));
    let xml = String::from_utf8(spec.to_xml(Some("rId1")).unwrap()).unwrap();
    assert!(xml.contains("<c:f>Sheet1!$A$2:$A$5</c:f>"), "{xml}");
    assert!(xml.contains("<c:f>Sheet1!$B$2:$B$5</c:f>"));
    assert!(xml.contains("<c:f>Sheet1!$C$2:$C$5</c:f>"));
    assert!(xml.contains(r#"<c:externalData r:id="rId1""#));
    let layout = spec.sheet_layout();
    assert!(layout.contains(&(0, 1, Cell::Text("East"))));
    assert!(layout.contains(&(1, 0, Cell::Text("Q1"))));
    assert!(layout.contains(&(4, 2, Cell::Number(128.0))));

    let scatter = Spec::new(Variant::new(Kind::Scatter))
        .with_series(Series::new("A", vec![10.0, 20.0]))
        .with_binding(Binding::sheet("Data Sheet"));
    let xml = String::from_utf8(scatter.to_xml(None).unwrap()).unwrap();
    assert!(
        xml.contains("<c:f>&apos;Data Sheet&apos;!$A$2:$A$3</c:f>"),
        "{xml}"
    );
    assert!(xml.contains("<c:f>&apos;Data Sheet&apos;!$B$2:$B$3</c:f>"));
    assert_eq!(scatter.x_values(), vec![1.0, 2.0]);
}

#[test]
fn invalid_specs_are_refused() {
    assert!(Spec::new(Variant::new(Kind::Column)).validate().is_err());
    let nan = Spec::new(Variant::new(Kind::Line)).with_series(Series::new("A", vec![f64::NAN]));
    assert!(nan.validate().is_err());
    let empty = Spec::new(Variant::new(Kind::Line)).with_series(Series::new("A", vec![]));
    assert!(empty.validate().is_err());
    assert!(!Variant::new(Kind::Radar).with_three_d(true).is_valid());
    assert!(
        !Variant::new(Kind::Pie)
            .with_grouping(Grouping::Stacked)
            .is_valid()
    );
}

#[test]
fn source_preserving_edits_change_title_legend_axes_and_data() {
    let spec = sales(Variant::new(Kind::Column))
        .with_binding(Binding::sheet("Sheet1"))
        .with_value_axis_title("Units");
    let snapshot = Package::from_xml(spec.to_xml(Some("rId1")).unwrap()).unwrap();
    let mut edit = snapshot.edit();
    edit.set_title(Some(TitleText::from_string("Revenue")))
        .unwrap();
    edit.set_legend(Some(LegendPosition::Top)).unwrap();
    edit.set_axis_title(CATEGORY_AXIS_ID, Some(TitleText::from_string("Quarter")))
        .unwrap();
    edit.set_axis_title(VALUE_AXIS_ID, None).unwrap();
    edit.set_series_title(1, Some(TitleText::from_string("South")))
        .unwrap();
    edit.set_series_values(0, &[1.0, 2.0, 3.0]).unwrap();
    edit.set_categories(&["A".into(), "B".into(), "C".into()])
        .unwrap();
    let after = edit.commit().unwrap().into_package();
    let chart = after.value();
    assert_eq!(text(chart.title.as_ref()).as_deref(), Some("Revenue"));
    assert_eq!(
        chart.legend.as_ref().map(|legend| legend.position),
        Some(LegendPosition::Top)
    );
    let series = &chart.plot_area.type_groups[0].common().series;
    assert_eq!(text(series[1].title.as_ref()).as_deref(), Some("South"));
    let values = series[0].values.as_ref().unwrap();
    assert_eq!(values.values, vec![1.0, 2.0, 3.0]);
    assert_eq!(
        values.source_ref.as_ref().unwrap().formula,
        "Sheet1!$B$2:$B$4"
    );
    let categories = series[1].categories.as_ref().unwrap();
    assert_eq!(categories.values, vec!["A", "B", "C"]);
    assert_eq!(
        categories.source_ref.as_ref().unwrap().formula,
        "Sheet1!$A$2:$A$4"
    );
    let xml = String::from_utf8(after.xml_bytes().to_vec()).unwrap();
    assert!(xml.contains("Quarter"));
    assert!(!xml.contains("Units"));

    let mut edit = after.edit();
    edit.set_legend(None).unwrap();
    let hidden = edit.commit().unwrap().into_package();
    assert!(hidden.value().legend.is_none());
    // An exact no-op keeps the bytes.
    let mut edit = hidden.edit();
    edit.set_legend(None).unwrap();
    assert!(!edit.commit().unwrap().diagnostics().changed());
}

#[test]
fn literal_data_is_rewritten_in_place() {
    let spec = sales(Variant::new(Kind::Line));
    let snapshot = Package::from_xml(spec.to_xml(None).unwrap()).unwrap();
    let mut edit = snapshot.edit();
    edit.set_series_values(1, &[7.0, 8.0]).unwrap();
    let after = edit.commit().unwrap().into_package();
    let series = &after.value().plot_area.type_groups[0].common().series;
    assert_eq!(series[1].values.as_ref().unwrap().values, vec![7.0, 8.0]);
    assert!(series[1].values.as_ref().unwrap().source_ref.is_none());
}

#[test]
fn column_names_follow_spreadsheet_order() {
    assert_eq!(column_name(0), "A");
    assert_eq!(column_name(25), "Z");
    assert_eq!(column_name(26), "AA");
    assert_eq!(column_name(701), "ZZ");
    assert_eq!(column_name(702), "AAA");
}

#[test]
fn a_range_origin_offsets_every_reference() {
    let spec = sales(Variant::new(Kind::Bar)).with_binding(Binding::Sheet {
        sheet: "Data".into(),
        row: 2,
        column: 3,
    });
    let xml = String::from_utf8(spec.to_xml(None).unwrap()).unwrap();
    assert!(xml.contains("<c:f>Data!$D$4:$D$7</c:f>"), "{xml}");
    assert!(xml.contains("<c:f>Data!$E$4:$E$7</c:f>"));
    let far = sales(Variant::new(Kind::Bar)).with_binding(Binding::Sheet {
        sheet: "Data".into(),
        row: 1_048_575,
        column: 0,
    });
    assert!(far.validate().is_err());
}

#[test]
fn appearance_options_reach_the_chart_part() {
    let appearance = Appearance {
        colors: vec!["4472C4".into(), "ed7d31".into()],
        labels: Some(Labels {
            value: true,
            position: Some(DataLabelPosition::OutsideEnd),
            ..Labels::default()
        }),
        gap_width: Some(80),
        overlap: Some(-20),
        legend_overlay: Some(true),
        axis_min: Some(0.0),
        axis_max: Some(200.0),
        major_unit: Some(50.0),
        gridlines: Some(Some(Line {
            color: "D9D9D9".into(),
            width: Some(0.5),
            dash: Some("dot".into()),
        })),
        minor_gridlines: Some(None),
        ..Appearance::default()
    };
    let spec = sales(Variant::new(Kind::Column)).with_appearance(appearance);
    let xml = String::from_utf8(spec.to_xml(None).unwrap()).unwrap();
    for needle in [
        r#"<a:srgbClr val="4472C4"/>"#,
        r#"<a:srgbClr val="ED7D31"/>"#,
        r#"<c:gapWidth val="80"/>"#,
        r#"<c:overlap val="-20"/>"#,
        r#"<c:dLblPos val="outEnd"/>"#,
        r#"<c:showVal val="1"/>"#,
        r#"<c:max val="200"/>"#,
        r#"<c:majorUnit val="50"/>"#,
        r#"<a:prstDash val="dot"/>"#,
        "<c:minorGridlines",
    ] {
        assert!(xml.contains(needle), "{needle} in {xml}");
    }
    let chart = reader::read(xml.as_bytes(), &crate::chart::Limits::DEFAULT).unwrap();
    assert!(chart.legend.unwrap().overlay);

    let pie = Spec::new(Variant::new(Kind::Doughnut))
        .with_series(Series::new("Share", vec![1.0, 2.0, 3.0]))
        .with_appearance(Appearance {
            colors: vec!["FF0000".into(), "00FF00".into(), "0000FF".into()],
            hole_size: Some(60),
            first_slice_angle: Some(90),
            explosion: Some(10),
            ..Appearance::default()
        });
    let xml = String::from_utf8(pie.to_xml(None).unwrap()).unwrap();
    assert!(xml.contains(r#"<c:holeSize val="60"/>"#), "{xml}");
    assert!(xml.contains(r#"<c:firstSliceAng val="90"/>"#));
    assert_eq!(xml.matches("<c:dPt>").count(), 3);

    for bad in [
        Appearance {
            colors: vec!["blue".into()],
            ..Appearance::default()
        },
        Appearance {
            hole_size: Some(95),
            ..Appearance::default()
        },
        Appearance {
            axis_min: Some(5.0),
            axis_max: Some(1.0),
            ..Appearance::default()
        },
        Appearance {
            labels: Some(Labels {
                value: true,
                position: Some(DataLabelPosition::Top),
                ..Labels::default()
            }),
            ..Appearance::default()
        },
    ] {
        assert!(
            sales(Variant::new(Kind::Column))
                .with_appearance(bad.clone())
                .validate()
                .is_err(),
            "{bad:?}"
        );
    }
}

#[test]
fn combo_charts_group_series_by_family() {
    let spec = Spec::new(Variant::new(Kind::Combo))
        .with_categories(["Q1", "Q2"])
        .with_series(Series::new("A", vec![1.0, 2.0]))
        .with_series(Series::new("B", vec![3.0, 4.0]))
        .with_series(Series::new("C", vec![5.0, 6.0]))
        .with_series_kinds(vec![Kind::Column, Kind::Column, Kind::Line])
        .with_appearance(Appearance {
            marker: Some(MarkerStyle::Diamond),
            marker_size: Some(7),
            line_width: Some(2.0),
            view_3d: Some((15, 20, Some(30))),
            ..Appearance::default()
        });
    let chart = reader::read(
        &spec.to_xml(None).unwrap()[..],
        &crate::chart::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(Variant::of(&chart), Some(Variant::new(Kind::Combo)));
    assert_eq!(chart.plot_area.type_groups.len(), 2);
    assert_eq!(chart.plot_area.type_groups[0].common().series.len(), 2);
    let line = &chart.plot_area.type_groups[1].common().series[0];
    assert_eq!(line.index, 2);
    assert_eq!(line.marker_symbol, Some(MarkerStyle::Diamond));
    assert_eq!(
        chart.view_3d.as_ref().and_then(|view| view.perspective),
        Some(30)
    );
    assert!(
        Spec::new(Variant::new(Kind::Combo))
            .with_series(Series::new("A", vec![1.0]))
            .validate()
            .is_err()
    );
    assert!(
        sales(Variant::new(Kind::Combo))
            .with_series_kinds(vec![Kind::Pie])
            .validate()
            .is_err()
    );
    assert!(
        sales(Variant::new(Kind::Line))
            .with_series_kinds(vec![Kind::Line])
            .validate()
            .is_err()
    );
}

#[test]
fn series_colors_override_the_palette_and_tint_whole_pies() {
    // A series color wins over its palette entry; the other series keep
    // theirs, and a series with neither keeps the application default.
    let spec = Spec::new(Variant::new(Kind::Line))
        .with_series(Series::new("A", vec![1.0, 2.0]).with_color("c00000"))
        .with_series(Series::new("B", vec![3.0, 4.0]))
        .with_series(Series::new("C", vec![5.0, 6.0]).with_color("8000FF00"))
        .with_appearance(Appearance {
            colors: vec!["111111".into(), "222222".into()],
            ..Appearance::default()
        });
    let xml = String::from_utf8(spec.to_xml(None).unwrap()).unwrap();
    assert!(xml.contains(r#"<a:srgbClr val="C00000"/>"#), "{xml}");
    assert!(!xml.contains(r#"val="111111""#));
    assert!(xml.contains(r#"<a:srgbClr val="222222"/>"#));
    assert!(xml.contains(r#"<a:srgbClr val="00FF00"><a:alpha val="50196"/></a:srgbClr>"#));
    assert_eq!(xml, String::from_utf8(spec.to_xml(None).unwrap()).unwrap());
    let chart = reader::read(xml.as_bytes(), &crate::chart::Limits::DEFAULT).unwrap();
    let series = &chart.plot_area.type_groups[0].common().series;
    assert!(
        series
            .iter()
            .all(|series| series.shape_properties.is_some())
    );

    // A single-series pie: the series color fills the series, the palette
    // still colors its points.
    let pie = Spec::new(Variant::new(Kind::Pie))
        .with_series(Series::new("Share", vec![1.0, 2.0]).with_color("4472C4"))
        .with_appearance(Appearance {
            colors: vec!["FF0000".into(), "00FF00".into()],
            ..Appearance::default()
        });
    let xml = String::from_utf8(pie.to_xml(None).unwrap()).unwrap();
    assert_eq!(xml.matches("<c:dPt>").count(), 2, "{xml}");
    assert!(xml.contains(r#"<a:srgbClr val="4472C4"/>"#));
    let chart = reader::read(xml.as_bytes(), &crate::chart::Limits::DEFAULT).unwrap();
    assert!(
        chart.plot_area.type_groups[0].common().series[0]
            .shape_properties
            .is_some()
    );

    let bad = Spec::new(Variant::new(Kind::Column))
        .with_series(Series::new("A", vec![1.0]).with_color("blue"));
    assert!(matches!(bad.validate(), Err(Error::Invalid(_))));
    assert!(bad.to_xml(None).is_err());
}

#[test]
fn every_variant_reads_back_as_its_authored_spec() {
    let kinds = [
        Kind::Column,
        Kind::Bar,
        Kind::Line,
        Kind::Pie,
        Kind::Doughnut,
        Kind::Area,
        Kind::Scatter,
        Kind::Bubble,
        Kind::Radar,
    ];
    for kind in kinds {
        for grouping in [
            Grouping::Standard,
            Grouping::Stacked,
            Grouping::PercentStacked,
        ] {
            for three_d in [false, true] {
                let variant = Variant::new(kind)
                    .with_grouping(grouping)
                    .with_three_d(three_d);
                if !variant.is_valid() {
                    continue;
                }
                let mut spec = sales(variant);
                if matches!(kind, Kind::Pie | Kind::Doughnut) {
                    spec.series.truncate(1);
                }
                for (binding, external) in [
                    (Binding::Literal, None),
                    (Binding::sheet("Sheet1"), Some("rId1")),
                ] {
                    let spec = spec.clone().with_binding(binding.clone());
                    let xml = spec.to_xml(external).unwrap();
                    let parsed = read(&xml, &crate::chart::Limits::DEFAULT).unwrap();
                    assert_eq!(parsed.spec.variant(), variant);
                    assert_eq!(parsed.spec.binding(), &binding, "{variant:?}");
                    assert_eq!(parsed.external_data.as_deref(), external);
                    assert!(
                        parsed.is_authored(&xml, &crate::chart::Limits::DEFAULT),
                        "{variant:?} {binding:?}"
                    );
                }
            }
        }
    }
}

#[test]
fn formatting_reads_back_into_the_spec() {
    let appearance = Appearance {
        labels: Some(Labels {
            value: true,
            category: true,
            position: Some(DataLabelPosition::OutsideEnd),
            ..Labels::default()
        }),
        gap_width: Some(80),
        overlap: Some(-20),
        axis_min: Some(0.0),
        axis_max: Some(200.0),
        major_unit: Some(50.0),
        legend_overlay: Some(true),
        ..Appearance::default()
    };
    let spec = sales(Variant::new(Kind::Column))
        .with_category_axis_title("Quarter")
        .with_value_axis_title("Revenue")
        .with_legend(Some(LegendPosition::Right))
        .with_appearance(appearance.clone());
    let mut spec = spec;
    spec.series[1].color = Some("ED7D31".into());
    let xml = spec.to_xml(None).unwrap();
    let parsed = read(&xml, &crate::chart::Limits::DEFAULT).unwrap();
    assert!(parsed.is_authored(&xml, &crate::chart::Limits::DEFAULT));
    assert_eq!(parsed.spec.appearance(), &appearance);
    assert_eq!(parsed.spec.series()[1].color.as_deref(), Some("ED7D31"));
    assert_eq!(parsed.spec.category_axis_title(), Some("Quarter"));
    assert_eq!(parsed.spec.value_axis_title(), Some("Revenue"));
    // A changed family rebuilds with the same data and formatting.
    let bar = parsed
        .spec
        .clone()
        .with_variant(Variant::new(Kind::Bar))
        .to_xml(None)
        .unwrap();
    let rebuilt = read(&bar, &crate::chart::Limits::DEFAULT).unwrap();
    assert_eq!(rebuilt.spec.variant().kind, Kind::Bar);
    assert_eq!(rebuilt.spec.series(), parsed.spec.series());

    let pie = Spec::new(Variant::new(Kind::Doughnut))
        .with_categories(["A", "B", "C"])
        .with_series(Series::new("S", vec![1.0, 2.0, 3.0]))
        .with_appearance(Appearance {
            colors: vec!["FF0000".into(), "00FF00".into(), "0000FF".into()],
            hole_size: Some(60),
            first_slice_angle: Some(90),
            explosion: Some(10),
            ..Appearance::default()
        });
    let xml = pie.to_xml(None).unwrap();
    let parsed = read(&xml, &crate::chart::Limits::DEFAULT).unwrap();
    assert!(parsed.is_authored(&xml, &crate::chart::Limits::DEFAULT));
    assert_eq!(parsed.spec.appearance().colors.len(), 3);

    let combo = sales(Variant::new(Kind::Combo)).with_series_kinds(vec![Kind::Column, Kind::Line]);
    let xml = combo.to_xml(None).unwrap();
    let parsed = read(&xml, &crate::chart::Limits::DEFAULT).unwrap();
    assert_eq!(parsed.spec.series_kinds(), [Kind::Column, Kind::Line]);
    assert!(parsed.is_authored(&xml, &crate::chart::Limits::DEFAULT));
}

#[test]
fn foreign_formatting_is_not_authored() {
    let spec = sales(Variant::new(Kind::Column));
    let xml = String::from_utf8(spec.to_xml(None).unwrap()).unwrap();
    // A rounded-corner chart has a field the spec cannot write.
    let foreign = xml.replace(
        r#"<c:roundedCorners val="0"/>"#,
        r#"<c:roundedCorners val="1"/>"#,
    );
    assert_ne!(foreign, xml);
    let parsed = read(foreign.as_bytes(), &crate::chart::Limits::DEFAULT).unwrap();
    assert!(!parsed.is_authored(foreign.as_bytes(), &crate::chart::Limits::DEFAULT));
}
