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

fn assert_unsupported(result: Result<(), WriteError>) {
    match result {
        Err(WriteError::Graph(litchi_ograph::Error::UnsupportedAuthoring { reason })) => {
            assert!(!reason.is_empty());
        },
        Err(error) => panic!("expected typed unsupported-authoring error, found {error}"),
        Ok(()) => panic!("incomplete chart authoring unexpectedly succeeded"),
    }
}

#[test]
fn valid_chart_requests_are_typed_atomic_refusals() {
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

    let before = write_to_bytes(&mut writer);
    for kind in [ChartKind::Bar, ChartKind::Line, ChartKind::Pie] {
        assert_unsupported(writer.add_chart(slide, 50, 50, 400, 240, chart(kind)));
    }
    let after = write_to_bytes(&mut writer);
    assert_eq!(after, before, "refusal must not mutate presentation state");

    let mut package = Package::from_reader(Cursor::new(after)).expect("open presentation");
    let presentation = package.presentation().expect("read presentation");
    let inventory = presentation.charts().expect("enumerate charts");
    assert_eq!(inventory.seen(), 0);
    assert_eq!(inventory.charts().count(), 0);
    assert_eq!(inventory.failures().count(), 0);
    assert!(
        presentation.slides().expect("read slides")[0]
            .shape_count()
            .expect("count shapes")
            >= 2
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
