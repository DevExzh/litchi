//! SpreadsheetDrawing regressions extracted from a real XLSX package.

use litchi_drawingml::chart::{reader as drawingml_reader, writer as drawingml_writer};
use litchi_opc::{OpcPackage, PackURI};
use litchi_spreadsheet_drawing::chart::anchor::write_all;
use litchi_spreadsheet_drawing::chart::{Anchor, read as read_chart, write};
use litchi_spreadsheet_drawing::shape::{Object, read as read_drawing};
use std::path::{Path, PathBuf};

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path)
}

fn part(path: &Path, name: &str) -> Vec<u8> {
    let package = OpcPackage::open(path).unwrap_or_else(|error| {
        panic!(
            "{} must remain a readable OPC package: {error}",
            path.display()
        )
    });
    let name = PackURI::new(name).unwrap_or_else(|error| panic!("invalid test part name: {error}"));
    package
        .get_part(&name)
        .unwrap_or_else(|error| panic!("{} must contain {name}: {error}", path.display()))
        .blob()
        .to_vec()
}

fn assert_minimal_xml(xml: &[u8]) {
    assert!(!xml.is_empty());
    assert!(
        !xml.iter().any(|byte| matches!(byte, b'\r' | b'\n' | b'\t')),
        "emitted XML must be one line without indentation: {}",
        String::from_utf8_lossy(xml)
    );
    assert!(
        !xml.windows(2).any(|bytes| bytes == b"> "),
        "emitted XML must not contain whitespace between elements: {}",
        String::from_utf8_lossy(xml)
    );
}

#[test]
fn extracts_real_shape_and_chart_parts_without_copying_fixture_xml() {
    let shape_package = fixture("ooxml/xlsx/fontSize.xlsx");
    let drawing_xml = part(&shape_package, "/xl/drawings/drawing1.xml");
    let objects =
        read_drawing(std::str::from_utf8(&drawing_xml).expect("OOXML drawing XML must be UTF-8"))
            .expect("real worksheet drawing must parse")
            .expect("fixture must contain a drawing object");
    assert!(matches!(objects[0].object, Object::Shape(_)));

    let chart_xml = part(
        &fixture("ooxml/xlsx/SimpleScatterChart.xlsx"),
        "/xl/charts/chart1.xml",
    );
    let chart = read_chart(chart_xml.as_slice(), Anchor::default())
        .expect("real worksheet chart must parse");
    assert_eq!(chart.series_count(), 1);

    let emitted = write(&chart.chart).expect("parsed worksheet chart must serialize");
    assert_minimal_xml(&emitted);

    let mut anchor_xml = String::new();
    write_all(&mut anchor_xml, &[chart], 1, 0)
        .expect("real worksheet chart must emit a worksheet anchor");
    assert_minimal_xml(anchor_xml.as_bytes());
    assert!(anchor_xml.contains("r:id=\"rId1\""));
}

#[test]
fn extracts_and_canonically_emits_powerpoint_chart_through_shared_drawingml() {
    let chart_xml = part(
        &fixture("ooxml/pptx/pie-chart.pptx"),
        "/ppt/charts/chart1.xml",
    );
    let chart = drawingml_reader::read(chart_xml.as_slice())
        .expect("real PowerPoint chart must parse through shared DrawingML");
    assert!(!chart.plot_area.type_groups.is_empty());

    let mut emitted = Vec::new();
    drawingml_writer::write(&mut emitted, &chart)
        .expect("real PowerPoint chart must serialize through shared DrawingML");
    assert_minimal_xml(&emitted);
    drawingml_reader::read(emitted.as_slice()).expect("canonical chart XML must reparse");
}
