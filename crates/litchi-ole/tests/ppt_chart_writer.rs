use litchi_ole::ppt::writer::{Chart, ChartKind, Hyperlink, PptWriter, Table};
use litchi_ole::ppt::{Package, PowerPointChartKind};
use litchi_ole::xls::{XlsChartCachedValue, XlsChartGroupKind, XlsChartLocation, XlsChartType};
use std::io::Cursor;

fn bar_chart() -> Chart {
    let mut chart = Chart::new(ChartKind::Bar);
    chart.set_title("Quarterly sales");
    chart.set_categories(["Q1", "Q2", "Q3", "Q4"]);
    chart
        .add_series(Some("2023"), vec![1.0, 2.0, 3.0, 4.0])
        .unwrap();
    chart
        .add_series(Some("2024"), vec![1.5, 2.5, 3.5, 4.5])
        .unwrap();
    chart
}

fn write_to_bytes(writer: &mut PptWriter) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn authored_bar_chart_round_trips_through_chart_inventory() {
    let mut writer = PptWriter::new();
    let slide = writer.add_slide().unwrap();
    // Coexistence: a text box and a table share the slide with the chart.
    writer
        .add_textbox(slide, 40, 10, 300, 30, "Sales report")
        .unwrap();
    let mut table = Table::new(2, 2).unwrap();
    table.set_cell_text(0, 0, "Quarter").unwrap();
    table.set_cell_text(0, 1, "Sales").unwrap();
    writer.add_table(slide, 40, 300, table).unwrap();
    writer
        .add_chart(slide, 50, 50, 400, 240, bar_chart())
        .unwrap();
    let bytes = write_to_bytes(&mut writer);

    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let inventory = presentation.charts().unwrap();

    assert_eq!(
        inventory.failures.len(),
        0,
        "failures: {:?}",
        inventory.failures
    );
    assert_eq!(inventory.charts.len(), 1);
    let chart = &inventory.charts[0];
    assert_eq!(chart.kind, PowerPointChartKind::ExcelChart);
    assert_eq!(chart.program_id.as_deref(), Some("Excel.Chart.8"));
    let frame = chart.frame.expect("authored chart frame is attributed");
    assert_eq!(frame.slide_number, 1);

    assert_eq!(chart.charts.len(), 1);
    let entry = &chart.charts[0];
    assert!(matches!(
        entry.location,
        XlsChartLocation::Embedded { sheet_index: 0, .. }
    ));
    assert!(matches!(
        entry.chart.chart_type(),
        XlsChartType::Single(XlsChartGroupKind::Bar { .. })
    ));
    assert_eq!(entry.chart.title.as_deref(), Some("Quarterly sales"));
    assert_eq!(entry.chart.series.len(), 2);
    assert_eq!(entry.chart.series[0].name.as_deref(), Some("2023"));
    assert_eq!(entry.chart.series[0].links.len(), 2);
    assert!(
        entry
            .chart
            .cached_values
            .iter()
            .any(|value| value.value == XlsChartCachedValue::Number(4.5))
    );
    assert!(
        entry
            .chart
            .cached_values
            .iter()
            .any(|value| value.value == XlsChartCachedValue::Text("Q3".to_string()))
    );

    // The table and text box survived alongside the chart frame.
    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), 1);
    assert!(slides[0].shape_count().unwrap() >= 3);
}

#[test]
fn authored_line_and_pie_charts_round_trip() {
    let mut writer = PptWriter::new();
    let slide = writer.add_slide().unwrap();

    let mut line = Chart::new(ChartKind::Line);
    line.set_title("Trend");
    line.set_categories(["Jan", "Feb"]);
    line.add_series(Some("visits"), vec![10.0, 12.5]).unwrap();
    writer.add_chart(slide, 50, 50, 300, 200, line).unwrap();

    let mut pie = Chart::new(ChartKind::Pie);
    pie.set_categories(["A", "B", "C"]);
    pie.add_series(None::<String>, vec![30.0, 45.0, 25.0])
        .unwrap();
    writer.add_chart(slide, 380, 50, 300, 200, pie).unwrap();

    let bytes = write_to_bytes(&mut writer);
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let inventory = presentation.charts().unwrap();

    assert_eq!(inventory.failures.len(), 0);
    assert_eq!(inventory.charts.len(), 2);
    assert!(matches!(
        inventory.charts[0].charts[0].chart.chart_type(),
        XlsChartType::Single(XlsChartGroupKind::Line { .. })
    ));
    assert_eq!(
        inventory.charts[0].charts[0].chart.title.as_deref(),
        Some("Trend")
    );
    assert!(matches!(
        inventory.charts[1].charts[0].chart.chart_type(),
        XlsChartType::Single(XlsChartGroupKind::Pie { .. })
    ));
    assert_eq!(inventory.charts[1].charts[0].chart.series.len(), 1);
    // Both frames are attributed to the same slide with distinct shape ids.
    let first = inventory.charts[0].frame.unwrap();
    let second = inventory.charts[1].frame.unwrap();
    assert_eq!((first.slide_number, second.slide_number), (1, 1));
    assert_ne!(first.shape_id, second.shape_id);
}

#[test]
fn authored_chart_coexists_with_hyperlinks() {
    let mut writer = PptWriter::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 40, 10, 300, 30, "linked")
        .unwrap();
    let hyperlink_id = writer.add_hyperlink(Hyperlink::url("https://example.com"));
    writer
        .set_last_shape_hyperlink(slide, hyperlink_id)
        .unwrap();
    writer
        .add_chart(slide, 50, 60, 320, 220, bar_chart())
        .unwrap();
    let bytes = write_to_bytes(&mut writer);

    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    // The hyperlink and the chart share one ExObjList with distinct IDs.
    let objects = presentation
        .ole_objects()
        .unwrap()
        .expect("chart object list");
    assert_eq!(objects.objects.len(), 1);
    let inventory = presentation.charts().unwrap();
    assert_eq!(inventory.failures.len(), 0);
    assert_eq!(inventory.charts.len(), 1);
    assert!(inventory.charts[0].object_id > hyperlink_id);
}

#[test]
fn authored_chart_survives_save_to_disk() {
    let mut writer = PptWriter::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_chart(slide, 50, 50, 400, 240, bar_chart())
        .unwrap();

    let path = std::env::temp_dir().join(format!(
        "litchi-ppt-chart-{}-{}.ppt",
        std::process::id(),
        "roundtrip"
    ));
    writer.save(&path).unwrap();

    let mut package = Package::open(&path).unwrap();
    let presentation = package.presentation().unwrap();
    let inventory = presentation.charts().unwrap();
    assert_eq!(inventory.failures.len(), 0);
    assert_eq!(inventory.charts.len(), 1);
    assert_eq!(
        inventory.charts[0].charts[0].chart.title.as_deref(),
        Some("Quarterly sales")
    );
    std::fs::remove_file(&path).ok();
}

#[test]
fn invalid_chart_and_slide_are_rejected() {
    let mut writer = PptWriter::new();
    let slide = writer.add_slide().unwrap();
    // No series at all.
    assert!(
        writer
            .add_chart(slide, 0, 0, 100, 100, Chart::new(ChartKind::Bar))
            .is_err()
    );
    // Non-positive frame dimensions.
    assert!(writer.add_chart(slide, 0, 0, 0, 100, bar_chart()).is_err());
    // Unknown slide index.
    assert!(writer.add_chart(9, 0, 0, 100, 100, bar_chart()).is_err());
}
