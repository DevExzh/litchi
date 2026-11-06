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
    assert_eq!(pivot.name, "[excel-pivot-controls.xlsb]Sheet2!PivotTable1");
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
    let bytes = native("excel-chart-families.xlsb");
    let p = litchi_xlsb::Package::from_bytes(&bytes).unwrap();
    let wb = p.workbook().unwrap();
    assert!(wb.drawings_parsed());
    assert_eq!(wb.sheet_drawings().len(), 18);
    for (i, drawing) in wb.sheet_drawings().iter().enumerate() {
        assert_eq!(drawing.sheet_index, i);
        assert_eq!(drawing.charts.len(), usize::from(i != 7));
        if i == 7 {
            continue;
        }
        assert_eq!(drawing.charts[0].rel_id, "rId1");
        assert_family(&drawing.charts[0].chart, i);
    }
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_chart_sheets_keep_saved_sales_and_pivot_chart_semantics() {
    for (file, name) in [
        ("excel-values-formulas-chart.xlsb", "Sales chart"),
        ("excel-pivot-controls.xlsb", "PivotChart007"),
    ] {
        let p = litchi_xlsb::Package::from_bytes(&native(file)).unwrap();
        let wb = p.workbook().unwrap();
        assert_eq!(wb.chart_sheets().len(), 1);
        assert_eq!(wb.chart_sheets()[0].0, 0);
        let sheet = wb.chart_sheet(0).unwrap();
        assert_eq!(sheet.name, name);
        assert_eq!(sheet.drawing_rel_id.as_deref(), Some("rId1"));
        let drawing = wb.sheet_drawing(0).unwrap();
        assert_eq!(drawing.charts.len(), 1);
        let chart = &drawing.charts[0].chart;
        if name == "Sales chart" {
            assert_sales(chart);
        } else {
            assert_pivot(chart);
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
fn native_cell_edit_keeps_chart_graph_and_saved_caches_inert() {
    use litchi_core::sheet::{Cell as _, CellValue};
    let bytes = native("excel-chart-families.xlsb");
    let p = litchi_xlsb::Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit_cell_values("Line_013").unwrap();
    edit.set_scalar(
        litchi_xlsb::cell_values::Reference::new(1, 1).unwrap(),
        litchi_xlsb::cell_values::Scalar::Number(13.),
    )
    .unwrap();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    updated.save(output.path()).unwrap();
    let after_bytes = std::fs::read(output.path()).unwrap();
    let reopened = litchi_xlsb::Package::from_bytes(&after_bytes).unwrap();
    let wb = reopened.workbook().unwrap();
    assert_eq!(
        wb.worksheet(0).unwrap().get_cell(1, 1).unwrap().value(),
        &CellValue::Float(13.)
    );
    // Cell editing does not evaluate the chart; its saved cache stays at 10.
    assert_family(&wb.sheet_drawing(0).unwrap().charts[0].chart, 0);
    let before = members(&bytes);
    let after = members(&after_bytes);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    let changed = before
        .iter()
        .filter(|(name, data)| after[*name] != **data)
        .map(|(name, _)| name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        changed,
        ["xl/worksheets/binaryIndex1.bin", "xl/worksheets/sheet1.bin"]
    );
    assert!(
        reopened
            .apply_cell_values(&patch.inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}
