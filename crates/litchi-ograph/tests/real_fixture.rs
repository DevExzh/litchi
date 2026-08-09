//! MS-OGRAPH chart discovery over real legacy Excel workbooks.

use litchi_cfb::OleFile;
use litchi_ograph::chart::{Chart, Context, Kind, Refs};
use std::io::Cursor;
use std::path::PathBuf;

fn workbook_fixture(name: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/xls")
        .join(name);
    let source = std::fs::read(&path)
        .unwrap_or_else(|error| panic!("{} must remain readable: {error}", path.display()));
    let mut ole = OleFile::open(Cursor::new(source)).unwrap_or_else(|error| {
        panic!(
            "{} must remain a valid CFB workbook: {error}",
            path.display()
        )
    });
    ole.open_stream(&["Workbook"])
        .unwrap_or_else(|error| panic!("{} must contain Workbook: {error}", path.display()))
}

#[test]
fn discovers_and_replays_chart_substreams_from_real_excel_workbooks() {
    for name in ["SimpleChart.xls", "WithThreeCharts.xls"] {
        let workbook = workbook_fixture(name);
        let charts = Refs::open_workbook(&workbook)
            .expect("real XLS workbook framing must be valid")
            .collect::<Result<Vec<_>, _>>()
            .unwrap_or_else(|error| panic!("{name} chart discovery failed: {error}"));
        assert!(
            !charts.is_empty(),
            "{name} must retain an Excel chart substream"
        );

        for chart in charts {
            assert_eq!(chart.kind(), Kind::Excel, "{name}");
            let parsed = Chart::parse(chart, Context::excel())
                .unwrap_or_else(|error| panic!("{name} semantic chart parse failed: {error}"));
            let replayed = parsed
                .encode()
                .unwrap_or_else(|error| panic!("{name} untouched chart replay failed: {error}"));
            assert_eq!(replayed.as_bytes(), chart.as_bytes(), "{name}");
        }
    }
}
