#![allow(
    clippy::unwrap_used,
    reason = "corpus tests panic on unexpected fixture/provenance drift"
)]

use litchi_odc::{ChartClassKind, FlatChart};
use std::path::{Path, PathBuf};

fn repository_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .unwrap()
}

fn embedded_flat_chart(relative: &str) -> Vec<u8> {
    let source = std::fs::read_to_string(repository_root().join(relative)).unwrap();
    let mime = "office:mimetype=\"application/vnd.oasis.opendocument.chart\"";
    let mime_at = source.find(mime).unwrap();
    let start = source[..mime_at].rfind("<office:document ").unwrap();
    let closing = "</office:document>";
    let end = source[mime_at..].find(closing).unwrap() + mime_at + closing.len();
    source[start..end].as_bytes().to_vec()
}

#[test]
fn vendored_libreoffice_35_chart_subdocument_opens_losslessly() {
    let bytes = embedded_flat_chart(
        "test-data/libreoffice-core/chart2/qa/extras/data/fods/stacked-column-chart.fods",
    );
    let chart = FlatChart::from_bytes(bytes.clone()).unwrap();
    assert_eq!(chart.as_bytes(), bytes);
    assert_eq!(
        chart.chart().chart_class().unwrap().kind(),
        ChartClassKind::Bar
    );
    assert!(chart.find_axis("primary-x").is_some());
    assert!(chart.plot_area().unwrap().series().count() >= 1);
}

#[test]
fn vendored_libreoffice_25_odf_14_chart_subdocument_opens_losslessly() {
    let bytes =
        embedded_flat_chart("test-data/libreoffice-core/sw/qa/core/doc/data/tdf171549.fodt");
    let chart = FlatChart::from_bytes(bytes.clone()).unwrap();
    assert_eq!(chart.as_bytes(), bytes);
    assert_eq!(
        chart.chart().chart_class().unwrap().kind(),
        ChartClassKind::Bar
    );
    assert!(chart.find_axis("primary-y").is_some());
    assert!(chart.plot_area().unwrap().series().count() >= 1);
}
