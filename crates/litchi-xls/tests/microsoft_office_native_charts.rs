//! Native BIFF chart records, inert caches and source-preserving edits.
use litchi_xls::{
    Package,
    chart::{
        CacheKind, CellRef, GroupKind, Kind, Limits, Location, Role, Selector, Source, Value,
        chart_area::Rect,
    },
};
use std::{collections::BTreeMap, io::Cursor};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = ole.open_stream(&names).unwrap();
            (path, data)
        })
        .collect()
}

fn records(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut result = Vec::new();
    let mut pos = 0;
    while pos < bytes.len() {
        let kind = u16::from_le_bytes(bytes[pos..pos + 2].try_into().unwrap());
        let len = usize::from(u16::from_le_bytes(
            bytes[pos + 2..pos + 4].try_into().unwrap(),
        ));
        result.push((kind, bytes[pos + 4..pos + 4 + len].to_vec()));
        pos += 4 + len;
    }
    result
}

fn assert_other_streams(before: &[u8], after: &[u8]) {
    let before = streams(before);
    let after = streams(after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (path, data) in &before {
        if path != &["Workbook"] {
            assert!(after[path] == *data, "changed {path:?}");
        }
    }
}

#[test]
fn native_radar_is_discoverable_with_three_series_and_biff_cell_links() {
    let bytes = native("excel-chart-families.xls");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit_charts().unwrap();
    let entries = edit.charts().unwrap().collect::<Vec<_>>();
    assert_eq!(entries.len(), 17);
    assert_eq!(
        entries
            .iter()
            .map(|c| c.location.sheet_index())
            .collect::<Vec<_>>(),
        (0..18).filter(|i| *i != 7).collect::<Vec<_>>()
    );
    let radar = entries
        .iter()
        .find(|c| c.location.sheet_index() == 3)
        .unwrap();
    assert_eq!(
        radar.location,
        Location::Embedded {
            sheet_index: 3,
            object_id: 1
        }
    );
    let chart = &radar.chart;
    assert_eq!(chart.kind(), Kind::Radar);
    assert_eq!(chart.groups.len(), 1);
    assert!(matches!(
        chart.groups[0].kind,
        GroupKind::Radar { filled: false, .. }
    ));
    assert_eq!(chart.axes.len(), 2);
    assert_eq!(chart.series.len(), 3);
    for (i, s) in chart.series.iter().enumerate() {
        assert_eq!(
            s.name.as_deref(),
            Some(["SERIES_A_013", "SERIES_B_013", "SERIES_C_013"][i])
        );
        assert_eq!((s.category_count, s.value_count, s.bubble_count), (5, 5, 0));
        for (role, row, last, col) in [
            (Role::Name, 0, 0, i as u16 + 1),
            (Role::Values, 1, 5, i as u16 + 1),
            (Role::Categories, 1, 5, 0),
        ] {
            let link = s.link(role).unwrap();
            assert_eq!(link.source, Source::Cells);
            assert_eq!(
                link.references,
                [CellRef {
                    extern_sheet_index: 9,
                    first_row: row,
                    last_row: last,
                    first_column: col,
                    last_column: col
                }]
            );
        }
    }
    assert_eq!(
        chart
            .inventory(Limits::default())
            .unwrap()
            .cell_reference_count,
        9
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_sales_and_pivot_chart_sheets_expose_saved_series_and_caches() {
    for file in [
        "excel-values-formulas-chart.xls",
        "excel-pivot-controls.xls",
    ] {
        let p = Package::from_bytes(&native(file)).unwrap();
        let mut edit = p.edit_charts().unwrap();
        let entries = edit.charts().unwrap().collect::<Vec<_>>();
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].location, Location::ChartSheet { sheet_index: 0 });
        let c = &entries[0].chart;
        assert_eq!(c.kind(), Kind::Bar);
        assert_eq!(c.axes.len(), 2);
        if file.contains("values") {
            assert_eq!(c.series.len(), 3);
            assert_eq!(c.cached_values.len(), 24);
            for (i, s) in c.series.iter().enumerate() {
                assert_eq!(
                    s.name.as_deref(),
                    Some(["Quantity", "Unit price", "Total"][i])
                );
                assert_eq!((s.category_count, s.value_count), (4, 4));
                assert_eq!(
                    s.link(Role::Values).unwrap().references,
                    [CellRef {
                        extern_sheet_index: 0,
                        first_row: 1,
                        last_row: 4,
                        first_column: i as u16 + 1,
                        last_column: i as u16 + 1
                    }]
                );
                for (point, value) in [
                    [3., 5., 2., 4.],
                    [1.25, 2.5, 4.75, 6.],
                    [3.75, 12.5, 9.5, 24.],
                ][i]
                    .iter()
                    .enumerate()
                {
                    assert!(c.cached_values.iter().any(|v| v.kind == CacheKind::Values
                        && usize::from(v.series) == i
                        && usize::from(v.point) == point
                        && v.value == Value::Number(*value)));
                    assert!(
                        c.cached_values
                            .iter()
                            .any(|v| v.kind == CacheKind::Categories
                                && usize::from(v.series) == i
                                && usize::from(v.point) == point
                                && v.value
                                    == Value::Text(
                                        ["Apples", "Pears", "Tea", "Coffee"][point].into()
                                    ))
                    );
                }
            }
        } else {
            // BIFF stores the grand-total point too; the OOXML companion cache omits it.
            assert_eq!(c.series.len(), 1);
            assert_eq!(c.series[0].name.as_deref(), Some("Column Labels Gadget"));
            assert_eq!(
                (c.series[0].category_count, c.series[0].value_count),
                (2, 2)
            );
            assert_eq!(
                c.cached_values
                    .iter()
                    .map(|v| (&v.kind, v.point, &v.value))
                    .collect::<Vec<_>>(),
                [
                    (&CacheKind::Categories, 0, &Value::Text("North".into())),
                    (
                        &CacheKind::Categories,
                        1,
                        &Value::Text("Grand Total".into())
                    ),
                    (&CacheKind::Values, 0, &Value::Number(9.5)),
                    (&CacheKind::Values, 1, &Value::Number(9.5)),
                ]
            );
        }
    }
}
#[test]
fn native_chart_sheet_cache_edit_preserves_unrelated_records_and_inverts() {
    let bytes = native("excel-values-formulas-chart.xls");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit_charts().unwrap();
    let before_chart = edit.charts().unwrap().next().unwrap().chart.clone();
    edit.set_cache(
        Selector::Sheet("Sales chart"),
        CacheKind::Values,
        0,
        0,
        0,
        Value::Number(13.),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::new().unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let mut read = reopened.edit_charts().unwrap();
    let mut expected = before_chart;
    expected
        .cached_values
        .iter_mut()
        .find(|v| v.kind == CacheKind::Values && v.series == 0 && v.point == 0)
        .unwrap()
        .value = Value::Number(13.);
    assert_eq!(read.charts().unwrap().next().unwrap().chart, expected);
    assert_other_streams(&bytes, &after);
    let before = streams(&bytes);
    let after = streams(&after);
    let key = vec!["Workbook".to_owned()];
    let before = records(&before[&key]);
    let after = records(&after[&key]);
    assert_eq!(before.len(), after.len());
    let changed = before
        .iter()
        .zip(&after)
        .filter(|(a, b)| a != b)
        .collect::<Vec<_>>();
    assert_eq!(changed.len(), 1);
    assert_eq!(changed[0].0.0, 0x0203); // cached NUMBER, not a worksheet formula or cell
    assert!(
        reopened
            .apply_charts(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_embedded_chart_geometry_edits_change_only_the_selected_record() {
    let bytes = native("excel-chart-families.xls");
    let p = Package::from_bytes(&bytes).unwrap();
    for (sheet_index, sheet) in [
        (0, "Line_013"),
        (1, "Area_013"),
        (2, "Pie_013"),
        (3, "Radar_013"),
        (4, "Surface_013"),
        (5, "Scatter_013"),
        (6, "Bubble_013"),
        (8, "Bar_013"),
        (9, "Bar3D_013"),
        (10, "Area3D_013"),
        (11, "Line3D_013"),
        (12, "Contour_013"),
        (13, "Doughnut_013"),
        (14, "OfPie_013"),
        (15, "Pie3D_013"),
        (16, "Stock_013"),
        (17, "Combo_013"),
    ] {
        let chart_index = if sheet_index < 7 {
            sheet_index
        } else {
            sheet_index - 1
        };
        let mut edit = p.edit_charts().unwrap();
        let before = edit.charts().unwrap().cloned().collect::<Vec<_>>();
        let original = &before[chart_index].chart;
        let rect = Rect {
            x: original.x,
            y: original.y,
            width: original.width + (1 << 16),
            height: original.height,
        };
        edit.set_chart_area(Selector::Embedded { sheet, index: 0 }, rect)
            .unwrap();
        let commit = edit.commit().unwrap();
        let after_bytes = commit.package().to_bytes().unwrap();
        let reopened = Package::from_bytes(&after_bytes).unwrap();
        let mut read = reopened.edit_charts().unwrap();
        let mut expected = before;
        expected[chart_index].chart.width = rect.width;
        assert_eq!(
            read.charts().unwrap().cloned().collect::<Vec<_>>(),
            expected
        );
        assert_other_streams(&bytes, &after_bytes);
        let before = streams(&bytes);
        let after = streams(&after_bytes);
        let key = vec!["Workbook".to_owned()];
        let before = records(&before[&key]);
        let after = records(&after[&key]);
        assert_eq!(before.len(), after.len());
        let changed = before
            .iter()
            .zip(&after)
            .filter(|(a, b)| a != b)
            .collect::<Vec<_>>();
        assert_eq!(changed.len(), 1);
        assert_eq!(changed[0].0.0, 0x1002);
        assert!(
            reopened
                .apply_charts(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == bytes
        );
    }
}

#[test]
fn native_of_pie_retains_automatic_split_and_series_references() {
    use litchi_ograph::record::of_pie;
    let bytes = native("excel-chart-families.xls");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit_charts().unwrap();
    let entry = edit
        .charts()
        .unwrap()
        .find(|c| c.location.sheet_index() == 14)
        .unwrap();
    let c = &entry.chart;
    assert_eq!(c.kind(), Kind::OfPie);
    assert_eq!(c.groups.len(), 1);
    let GroupKind::OfPie { options } = c.groups[0].kind else {
        panic!("missing BopPop settings")
    };
    assert_eq!(options.kind(), of_pie::Kind::Pie);
    assert_eq!(options.split(), of_pie::Split::Automatic);
    assert_eq!(
        (
            options.secondary_size(),
            options.gap(),
            options.has_shadow()
        ),
        (75, 150, false)
    );
    assert_eq!(c.axes.len(), 0);
    assert_eq!(c.series.len(), 1);
    assert_eq!(c.series[0].name.as_deref(), Some("SERIES_A_013"));
    assert_eq!(
        (c.series[0].category_count, c.series[0].value_count),
        (5, 5)
    );
    for (role, first, last, col) in [
        (Role::Name, 0, 0, 1),
        (Role::Categories, 1, 5, 0),
        (Role::Values, 1, 5, 1),
    ] {
        assert_eq!(
            c.series[0].link(role).unwrap().references,
            [CellRef {
                extern_sheet_index: 15,
                first_row: first,
                last_row: last,
                first_column: col,
                last_column: col
            }]
        );
    }
    assert_eq!(c.groups[0].lines.len(), 1);
    assert_eq!(
        c.groups[0].lines[0].kind,
        litchi_ograph::record::line::Kind::Series
    );
    assert!(
        !c.unknown_records
            .iter()
            .any(|r| r.record_type == of_pie::Options::KIND.get())
    );
    assert!(p.to_bytes().unwrap() == bytes);
}

#[test]
fn native_stock_keeps_numeric_date_categories_and_high_low_lines() {
    use litchi_core::sheet::CellValue;
    use litchi_xls::chart::DataKind;
    let p = Package::from_bytes(&native("excel-chart-families.xls")).unwrap();
    let mut edit = p.edit_charts().unwrap();
    let entry = edit
        .charts()
        .unwrap()
        .find(|c| c.location.sheet_index() == 16)
        .unwrap();
    let c = &entry.chart;
    // MS-XLS encodes this profile with Line + CrtLine(HighLow), not a
    // distinct Stock record. High-low lines alone also occur on line charts.
    assert_eq!(c.kind(), Kind::Line);
    assert_eq!(c.groups.len(), 1);
    assert_eq!(c.groups[0].kind, GroupKind::Line { flags: 0 });
    assert!(c.groups[0].drop_bars.is_empty());
    assert_eq!(c.groups[0].lines.len(), 1);
    let line = &c.groups[0].lines[0];
    assert_eq!(line.kind, litchi_ograph::record::line::Kind::HighLow);
    assert_eq!(line.format.color, [51, 51, 51, 0]);
    assert_eq!((line.format.pattern, line.format.weight), (0, -1));
    assert_eq!(c.axes.len(), 2);
    assert_eq!(c.series.len(), 3);
    let workbook = p.workbook();
    let sheet = workbook.worksheet("Stock_013").unwrap().unwrap();
    for (i, s) in c.series.iter().enumerate() {
        assert_eq!(
            s.name.as_deref(),
            Some(["High_013", "Low_013", "Close_013"][i])
        );
        assert_eq!(s.category_kind, DataKind::Numeric);
        assert_eq!((s.category_count, s.value_count), (5, 5));
        let categories = s.link(Role::Categories).unwrap();
        assert_eq!(categories.number_format, 14);
        assert_eq!(
            categories.references,
            [CellRef {
                extern_sheet_index: 16,
                first_row: 1,
                last_row: 5,
                first_column: 0,
                last_column: 0
            }]
        );
        assert_eq!(
            s.link(Role::Values).unwrap().references,
            [CellRef {
                extern_sheet_index: 16,
                first_row: 1,
                last_row: 5,
                first_column: i as u16 + 1,
                last_column: i as u16 + 1
            }]
        );
        for (point, value) in [
            [15., 16., 18., 17., 20.],
            [10., 11., 13., 12., 14.],
            [12., 14., 15., 16., 18.],
        ][i]
            .iter()
            .enumerate()
        {
            let address =
                litchi_xls::cell_values::Reference::new(point as u32 + 1, i as u32 + 1).unwrap();
            assert_eq!(
                sheet.cell_value(address).unwrap(),
                Some(CellValue::Float(*value))
            );
            let date = litchi_xls::cell_values::Reference::new(point as u32 + 1, 0).unwrap();
            assert_eq!(
                sheet.cell_value(date).unwrap(),
                Some(CellValue::DateTime(46296. + point as f64))
            );
        }
    }
}
