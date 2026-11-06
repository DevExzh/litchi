//! Native Excel chart metadata; caches are saved values, not recalculation or rendering.
use litchi_drawingml::chart::{Chart, NumericData, TitleText, TypeGroup};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn group_name(group: &TypeGroup) -> &'static str {
    use TypeGroup::*;
    match group {
        Area(_) => "area",
        Area3D(_) => "area3D",
        Bar(_) => "bar",
        Bar3D(_) => "bar3D",
        Bubble(_) => "bubble",
        Doughnut(_) => "doughnut",
        Line(_) => "line",
        Line3D(_) => "line3D",
        OfPie(_) => "ofPie",
        Pie(_) => "pie",
        Pie3D(_) => "pie3D",
        Radar(_) => "radar",
        Scatter(_) => "scatter",
        Stock(_) => "stock",
        Surface(_) => "surface",
        Surface3D(_) => "surface3D",
    }
}
const SHEETS: [&str; 18] = [
    "Line_013",
    "Area_013",
    "Pie_013",
    "Radar_013",
    "Surface_013",
    "Scatter_013",
    "Bubble_013",
    "DrawingSVG_013",
    "Bar_013",
    "Bar3D_013",
    "Area3D_013",
    "Line3D_013",
    "Contour_013",
    "Doughnut_013",
    "OfPie_013",
    "Pie3D_013",
    "Stock_013",
    "Combo_013",
];
fn numeric(data: &Option<NumericData>, formula: &str, values: &[f64]) {
    let data = data.as_ref().unwrap();
    assert_eq!(data.source_ref.as_ref().unwrap().formula, formula);
    assert_eq!(data.values, values);
    assert_eq!(data.format_code.as_deref(), Some("General"));
}
fn assert_family(chart: &Chart, index: usize) {
    let groups: &[&str] = match index {
        0 => &["line"],
        1 => &["area"],
        2 => &["pie"],
        3 => &["radar"],
        4 => &["surface3D"],
        5 => &["scatter"],
        6 => &["bubble"],
        8 => &["bar"],
        9 => &["bar3D"],
        10 => &["area3D"],
        11 => &["line3D"],
        12 => &["surface"],
        13 => &["doughnut"],
        14 => &["ofPie"],
        15 => &["pie3D"],
        16 => &["stock"],
        17 => &["bar", "line"],
        _ => unreachable!(),
    };
    assert_eq!(
        chart
            .plot_area
            .type_groups
            .iter()
            .map(group_name)
            .collect::<Vec<_>>(),
        groups
    );
    let axes = match index {
        2 | 13..=15 => 0,
        4 | 10..=12 => 3,
        17 => 4,
        _ => 2,
    };
    assert_eq!(chart.plot_area.axes.len(), axes);
    assert!(chart.pivot_source.is_none());
    let series = chart
        .plot_area
        .type_groups
        .iter()
        .flat_map(|g| &g.common().series)
        .collect::<Vec<_>>();
    assert_eq!(
        series.len(),
        if matches!(index, 2 | 5 | 6 | 13..=15) {
            1
        } else {
            3
        }
    );
    let sheet = SHEETS[index];
    for (i, s) in series.iter().enumerate() {
        assert_eq!((s.index, s.order), (i as u32, i as u32));
        let column = ['B', 'C', 'D'][i];
        let Some(TitleText::Reference(title)) = &s.title else {
            panic!("native series reference missing")
        };
        assert_eq!(title.formula, format!("{sheet}!${column}$1"));
        if index == 5 || index == 6 {
            numeric(
                &s.x_values,
                &format!("{sheet}!$A$2:$A$6"),
                &[1., 2., 3., 4., 5.],
            );
            numeric(
                &s.y_values,
                &format!("{sheet}!$B$2:$B$6"),
                &[2., 4., 3., 8., 7.],
            );
            assert!(s.categories.is_none() && s.values.is_none());
            if index == 6 {
                numeric(
                    &s.bubble_sizes,
                    &format!("{sheet}!$C$2:$C$6"),
                    &[5., 10., 7., 14., 11.],
                );
            } else {
                assert!(s.bubble_sizes.is_none());
            }
        } else {
            let categories = s.categories.as_ref().unwrap();
            assert_eq!(
                categories.source_ref.as_ref().unwrap().formula,
                format!("{sheet}!$A$2:$A$6")
            );
            let expected = if index == 16 {
                ["46296", "46297", "46298", "46299", "46300"]
            } else {
                [
                    "ALPHA_013",
                    "BETA_013",
                    "GAMMA_013",
                    "DELTA_013",
                    "EPSILON_013",
                ]
            };
            assert_eq!(categories.values, expected);
            let values = if index == 16 {
                [
                    [15., 16., 18., 17., 20.],
                    [10., 11., 13., 12., 14.],
                    [12., 14., 15., 16., 18.],
                ]
            } else {
                [
                    [10., 15., 12., 20., 17.],
                    [14., 11., 18., 16., 21.],
                    [8., 13., 16., 22., 19.],
                ]
            };
            numeric(
                &s.values,
                &format!("{sheet}!${column}$2:${column}$6"),
                &values[i],
            );
            assert!(s.x_values.is_none() && s.y_values.is_none() && s.bubble_sizes.is_none());
        }
        assert_eq!(
            s.error_bars.len(),
            match index {
                0 => 1,
                5 => 2,
                _ => 0,
            }
        );
        assert_eq!(
            s.trendlines.len(),
            usize::from((index == 0 && i == 0) || index == 5)
        );
    }
    use litchi_drawingml::chart::series::{
        ErrorBarDirection, ErrorBarType, ErrorBarValueType, TrendlineType,
    };
    for s in &series {
        for error in &s.error_bars {
            assert_eq!(error.error_type, ErrorBarType::Both);
            if index == 0 {
                assert_eq!(error.value_type, ErrorBarValueType::Percentage);
                assert_eq!(error.value, Some(5.));
                assert_eq!(error.direction, ErrorBarDirection::Y);
            } else if error.direction == ErrorBarDirection::X {
                assert_eq!(error.value_type, ErrorBarValueType::Fixed);
                assert_eq!(error.value, Some(1.));
            } else {
                assert_eq!(error.value_type, ErrorBarValueType::Custom);
                let plus = error.plus_values.as_ref().unwrap();
                let minus = error.minus_values.as_ref().unwrap();
                assert!(plus.source_ref.is_none() && minus.source_ref.is_none());
                assert_eq!(plus.values, [0.5, 1., 1.5, 2., 2.5]);
                assert_eq!(minus.values, [0.25, 0.5, 0.75, 1., 1.25]);
            }
        }
        for trend in &s.trendlines {
            assert_eq!(trend.trendline_type, TrendlineType::Linear);
            assert_eq!(trend.display_equation, index == 5);
            assert_eq!(trend.display_r_squared, index == 5);
        }
    }
    if index == 17 {
        assert_eq!(chart.plot_area.type_groups[0].common().series.len(), 2);
        assert_eq!(chart.plot_area.type_groups[1].common().series[0].index, 2);
    }
}
fn assert_sales(chart: &Chart) {
    assert!(chart.pivot_source.is_none());
    assert_eq!(chart.plot_area.type_groups.len(), 1);
    assert_eq!(group_name(&chart.plot_area.type_groups[0]), "bar");
    assert_eq!(chart.plot_area.axes.len(), 2);
    let series = &chart.plot_area.type_groups[0].common().series;
    assert_eq!(series.len(), 3);
    for (i, s) in series.iter().enumerate() {
        let col = ['B', 'C', 'D'][i];
        assert_eq!(
            s.categories.as_ref().unwrap().values,
            ["Apples", "Pears", "Tea", "Coffee"]
        );
        assert_eq!(
            s.categories
                .as_ref()
                .unwrap()
                .source_ref
                .as_ref()
                .unwrap()
                .formula,
            "Sheet1!$A$2:$A$5"
        );
        numeric(
            &s.values,
            &format!("Sheet1!${col}$2:${col}$5"),
            &[
                [3., 5., 2., 4.],
                [1.25, 2.5, 4.75, 6.],
                [3.75, 12.5, 9.5, 24.],
            ][i],
        );
    }
}
fn assert_pivot(chart: &Chart) {
    let pivot = chart.pivot_source.as_ref().unwrap();
    assert_eq!(pivot.name, "[excel-pivot-controls.xlsx]Sheet2!PivotTable1");
    assert_eq!(pivot.format_id, 1);
    assert_eq!(chart.plot_area.type_groups.len(), 1);
    assert_eq!(group_name(&chart.plot_area.type_groups[0]), "bar");
    let series = &chart.plot_area.type_groups[0].common().series;
    assert_eq!(series.len(), 1);
    assert_eq!(series[0].categories.as_ref().unwrap().values, ["North"]);
    assert_eq!(
        series[0]
            .categories
            .as_ref()
            .unwrap()
            .source_ref
            .as_ref()
            .unwrap()
            .formula,
        "Sheet2!$A$5:$A$6"
    );
    numeric(&series[0].values, "Sheet2!$B$5:$B$6", &[9.5]);
}

#[test]
fn native_seventeen_chart_families_keep_formula_cache_and_axis_semantics() {
    let bytes = native("excel-chart-families.xlsx");
    let p = litchi_xlsx::Package::from_bytes(&bytes).unwrap();
    let wb = p.workbook();
    assert_eq!(
        wb.sheets().map(|s| s.name().to_owned()).collect::<Vec<_>>(),
        SHEETS
    );
    for i in 0..18 {
        assert_eq!(p.chart_count(i).unwrap(), usize::from(i != 7));
        if i == 7 {
            continue;
        }
        let chart = litchi_xlsx::chart::read(
            p.chart_xml(i, 0).unwrap().unwrap(),
            p.chart_anchor(i, 0).unwrap().unwrap(),
        )
        .unwrap();
        assert_family(&chart.chart, i);
    }
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_chart_sheets_and_pivot_binding_resolve_their_resource_graphs() {
    for (file, name) in [
        ("excel-values-formulas-chart.xlsx", "Sales chart"),
        ("excel-pivot-controls.xlsx", "PivotChart007"),
    ] {
        let opc = litchi_xlsx::Package::from_bytes(&native(file))
            .unwrap()
            .into_plain_opc();
        let wb = litchi_opc::PackURI::new("/xl/workbook.xml").unwrap();
        let sheet = litchi_xlsx::chart_sheet::package::load_chartsheet(&opc, &wb, "rId1").unwrap();
        assert_eq!(sheet.entry.name, name);
        assert_eq!(sheet.entry.part_name, "/xl/chartsheets/sheet1.xml");
        assert_eq!(sheet.drawing.charts.len(), 1);
        let resource = &sheet.drawing.charts[0];
        assert_eq!(resource.part_name, "/xl/charts/chart1.xml");
        let chart = litchi_spreadsheet_drawing::chart::decode(&resource.data).unwrap();
        if name == "Sales chart" {
            assert_sales(&chart);
        } else {
            assert_pivot(&chart);
            let pivots = litchi_xlsx::pivot::chart::load(&opc).unwrap();
            assert_eq!(pivots.len(), 1);
            assert_eq!(pivots[0].sheet_name, name);
            assert_eq!(
                pivots[0].sheet_kind,
                litchi_xlsx::pivot::chart::SheetKind::Chartsheet
            );
            assert_eq!(pivots[0].charts.len(), 1);
            assert_eq!(
                pivots[0].charts[0].source.name,
                "[excel-pivot-controls.xlsx]Sheet2!PivotTable1"
            );
            assert_eq!(pivots[0].charts[0].pivot_table.name, "PivotTable1");
        }
    }
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

#[test]
fn native_embedded_chart_cache_update_preserves_companions_and_inverts() {
    let bytes = native("excel-chart-families.xlsx");
    let p = litchi_xlsx::Package::from_bytes(&bytes).unwrap();
    let xml = std::str::from_utf8(p.chart_xml(0, 0).unwrap().unwrap()).unwrap();
    assert_eq!(xml.matches("<c:v>10</c:v>").count(), 1);
    let changed = xml
        .replacen("<c:v>10</c:v>", "<c:v>13</c:v>", 1)
        .into_bytes();
    let mut edit = p.edit_charts().unwrap();
    assert!(edit.set(0, 0, changed.clone()).unwrap());
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::new().unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let reopened = litchi_xlsx::Package::from_bytes(&after).unwrap();
    assert_eq!(reopened.chart_xml(0, 0).unwrap().unwrap(), changed);
    let c =
        litchi_xlsx::chart::read(&changed, reopened.chart_anchor(0, 0).unwrap().unwrap()).unwrap();
    assert_eq!(
        c.chart.plot_area.type_groups[0].common().series[0]
            .values
            .as_ref()
            .unwrap()
            .values,
        [13., 15., 12., 20., 17.]
    );
    let before = members(&bytes);
    let after = members(&after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    let differences = before
        .iter()
        .filter(|(name, data)| after[*name] != **data)
        .map(|(name, _)| name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(differences, ["xl/charts/chart1.xml"]);
    let inverse = reopened.apply_charts(&commit.patch().inverse()).unwrap();
    let restored = members(&inverse.package().to_bytes().unwrap());
    let inverse_differences = before
        .iter()
        .filter(|(name, data)| restored[*name] != **data)
        .map(|(name, _)| name.as_str())
        .collect::<Vec<_>>();
    assert!(
        inverse_differences.is_empty(),
        "inverse changed {inverse_differences:?}"
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_implicit_chart_companions_validate_without_weakening_other_relationships() {
    use litchi_xlsx::chart_sheet::{Conformance, package::*};
    let opc = litchi_xlsx::Package::from_bytes(&native("excel-values-formulas-chart.xlsx"))
        .unwrap()
        .into_plain_opc();
    let wb = litchi_opc::PackURI::new("/xl/workbook.xml").unwrap();
    let native = load_chartsheet(&opc, &wb, "rId1").unwrap();
    validate_package(&native, Conformance::Transitional).unwrap();
    let ChartResourceKind::ClassicWithRelationships {
        outbound_resources, ..
    } = &native.drawing.charts[0].kind
    else {
        panic!("native style resources missing")
    };
    assert_eq!(outbound_resources.len(), 2);
    for (i, resource) in outbound_resources.iter().enumerate() {
        let ChartOutboundResource::Opaque(resource) = resource else {
            panic!("expected inert companion")
        };
        assert_eq!(resource.relationship_id, format!("rId{}", i + 1));
        assert_eq!(
            resource.part_name,
            ["/xl/charts/style1.xml", "/xl/charts/colors1.xml"][i]
        );
        assert_eq!(
            resource.data,
            opc.get_part(&litchi_opc::PackURI::new(&resource.part_name).unwrap())
                .unwrap()
                .blob()
        );
    }
    for case in 0..5 {
        let mut changed = native.clone();
        let ChartResourceKind::ClassicWithRelationships {
            outbound_resources, ..
        } = &mut changed.drawing.charts[0].kind
        else {
            unreachable!()
        };
        let ChartOutboundResource::Opaque(resource) = &mut outbound_resources[0] else {
            unreachable!()
        };
        match case {
            0 => resource.relationship_type = "urn:unreferenced".into(),
            1 => resource.content_type = "application/xml".into(),
            2 => resource.data = b"<wrong/>".to_vec(),
            3 => resource.part_name = "/xl/worksheets/style.xml".into(),
            _ => {
                outbound_resources[0] = ChartOutboundResource::External(ChartExternalResource {
                    relationship_id: resource.relationship_id.clone(),
                    relationship_type: resource.relationship_type.clone(),
                    target: "https://example.invalid/style.xml".into(),
                });
            },
        }
        assert!(
            validate_package(&changed, Conformance::Transitional).is_err(),
            "case {case}"
        );
    }
    let mut aliases = native;
    let ChartResourceKind::ClassicWithRelationships {
        outbound_resources, ..
    } = &mut aliases.drawing.charts[0].kind
    else {
        unreachable!()
    };
    for r in outbound_resources {
        let ChartOutboundResource::Opaque(r) = r else {
            unreachable!()
        };
        r.relationship_type = r.relationship_type.replace("/2011/", "/2012/");
    }
    validate_package(&aliases, Conformance::Transitional).unwrap();
}
