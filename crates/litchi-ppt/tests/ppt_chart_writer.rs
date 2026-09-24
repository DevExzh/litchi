#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::io::Cursor;

use litchi_ppt::Package;
use litchi_ppt::writer::{Chart, ChartKind, Hyperlink, Table, WriteError, Writer};

fn chart(kind: ChartKind) -> Chart {
    let mut chart = Chart::new(kind);
    chart.set_title("Quarterly sales");
    chart.set_categories(["Q1", "Q2", "Q3", "Q4"]);
    chart
        .add_series(Some("2024"), vec![1.5, 2.5, 3.5, 4.5])
        .expect("valid series");
    chart
}

fn write_to_bytes(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write presentation");
    output.into_inner()
}

#[test]
fn valid_chart_requests_write_graph_objects_and_reopen() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().expect("add slide");
    writer
        .add_textbox(slide, 40, 10, 300, 30, "Sales report")
        .expect("add text box");
    let mut table = Table::new(2, 2).expect("create table");
    table
        .set_cell_text(0, 0, "Quarter")
        .expect("set table cell");
    writer.add_table(slide, 40, 300, table).expect("add table");
    let link = writer.add_hyperlink(Hyperlink::url("https://example.com"));
    writer
        .set_last_shape_hyperlink(slide, link)
        .expect("link existing shape");

    for kind in [ChartKind::Bar, ChartKind::Line, ChartKind::Pie] {
        writer
            .add_chart(slide, 50, 50, 400, 240, chart(kind))
            .expect("standalone Graph chart");
    }
    let after = write_to_bytes(&mut writer);

    let mut package = Package::from_reader(Cursor::new(after)).expect("open presentation");
    let presentation = package.presentation().expect("read presentation");
    let inventory = presentation.charts().expect("enumerate charts");
    assert_eq!(inventory.seen(), 3);
    assert_eq!(inventory.charts().count(), 3);
    assert_eq!(inventory.failures().count(), 0);
    for chart in inventory.charts() {
        assert_eq!(chart.kind(), litchi_ppt::chart::Kind::Graph);
        assert_eq!(chart.info().program(), Some("MSGraph.Chart"));
        assert!(chart.info().frame().is_some());
        let litchi_ppt::chart::Chart::Graph(graph) = chart else {
            panic!("fresh writer chart must use Graph subtype");
        };
        assert_eq!(graph.book().len(), 1);
        let chart = graph
            .book()
            .charts()
            .next()
            .expect("one Graph chart")
            .expect("valid Graph chart");
        assert_eq!(chart.kind(), litchi_ograph::chart::Kind::Graph);
        let semantic = graph.semantic_chart().expect("semantic Graph chart");
        assert_eq!(semantic.title(), Some("Quarterly sales"));
        assert_eq!(semantic.series().len(), 1);
        assert_eq!(semantic.caches().len(), 9);
    }
    assert!(
        presentation.slides().expect("read slides")[0]
            .shape_count()
            .expect("count shapes")
            >= 5
    );
}

#[test]
fn malformed_chart_requests_still_report_input_errors() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().expect("add slide");

    assert!(matches!(
        writer.add_chart(slide, 0, 0, 100, 100, Chart::new(ChartKind::Bar)),
        Err(WriteError::InvalidData(_))
    ));
    assert!(matches!(
        writer.add_chart(slide, 0, 0, 0, 100, chart(ChartKind::Bar)),
        Err(WriteError::InvalidData(_))
    ));
    assert!(matches!(
        writer.add_chart(slide, i32::MAX, 0, 1, 100, chart(ChartKind::Bar)),
        Err(WriteError::InvalidData(_))
    ));
    assert!(matches!(
        writer.add_chart(9, 0, 0, 100, 100, chart(ChartKind::Bar)),
        Err(WriteError::InvalidData(_))
    ));
}

#[test]
fn chart_builder_rejects_invalid_series_before_authoring() {
    let mut chart = Chart::new(ChartKind::Line);
    assert!(matches!(
        chart.add_series(None::<String>, Vec::new()),
        Err(WriteError::InvalidData(_))
    ));
    assert!(matches!(
        chart.add_series(None::<String>, vec![f64::NAN]),
        Err(WriteError::InvalidData(_))
    ));
}

#[test]
fn graph_cache_limit_refusal_is_atomic_before_chart_storage() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().expect("add slide");
    let mut chart = Chart::new(ChartKind::Line);
    chart.set_categories((0..127).map(|index| format!("C{index}")));
    for _ in 0..255 {
        chart
            .add_series(None::<String>, vec![1.0; 127])
            .expect("bounded series");
    }

    assert!(matches!(
        writer.add_chart(slide, 0, 0, 100, 100, chart),
        Err(WriteError::Graph(litchi_ograph::Error::LimitExceeded {
            resource: "cached value count",
            ..
        }))
    ));
    let bytes = write_to_bytes(&mut writer);
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("open presentation");
    let inventory = package
        .presentation()
        .expect("presentation")
        .charts()
        .expect("chart inventory");
    assert!(inventory.is_empty());
}
