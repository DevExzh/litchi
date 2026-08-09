#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Integration tests for PPT table authoring: tables written by `Writer`
//! must round-trip through the table extraction APIs after save/reopen.

use litchi_ppt::Package;
use litchi_ppt::writer::{Table, Writer};
use std::io::Cursor;

fn write(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn reopen(bytes: Vec<u8>) -> Package<Cursor<Vec<u8>>> {
    Package::from_reader(Cursor::new(bytes)).unwrap()
}

#[test]
fn authored_2x3_table_round_trips_with_cell_text() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    let mut table = Table::new(2, 3).unwrap();
    let texts = [["A1", "B1", "C1"], ["A2", "B2", "C2"]];
    for (row, row_texts) in texts.iter().enumerate() {
        for (column, text) in row_texts.iter().enumerate() {
            table.set_cell_text(row, column, *text).unwrap();
        }
    }
    writer.add_table(slide, 72, 72, table).unwrap();

    let mut package = reopen(write(&mut writer));
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), 1);

    let shapes = slides[0].shapes().unwrap();
    assert_eq!(shapes.len(), 1);
    let read_table = shapes[0].as_table().expect("table group");
    assert_eq!(read_table.rows(), 2);
    assert_eq!(read_table.columns(), 3);
    for (row, row_texts) in texts.iter().enumerate() {
        for (column, text) in row_texts.iter().enumerate() {
            assert_eq!(read_table.cell(row, column), Some(*text));
        }
    }
}

#[test]
fn table_coexists_with_other_shapes() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    writer.add_rectangle(slide, 10, 10, 100, 50).unwrap();
    writer
        .add_textbox(slide, 10, 300, 300, 50, "caption")
        .unwrap();

    let mut table = Table::new(2, 2).unwrap();
    table.set_cell_text(0, 0, "x").unwrap();
    table.set_cell_text(1, 1, "y").unwrap();
    writer.add_table(slide, 50, 100, table).unwrap();

    let mut package = reopen(write(&mut writer));
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let shapes = slides[0].shapes().unwrap();

    let tables: Vec<_> = shapes.iter().filter_map(|shape| shape.as_table()).collect();
    assert_eq!(tables.len(), 1);
    assert_eq!((tables[0].rows(), tables[0].columns()), (2, 2));
    assert_eq!(tables[0].cell(0, 0), Some("x"));
    assert_eq!(tables[0].cell(1, 1), Some("y"));

    // The non-table shapes survive alongside the table.
    let texts: Vec<String> = shapes
        .iter()
        .filter_map(|shape| shape.text().ok())
        .collect();
    assert!(texts.iter().any(|text| text.contains("caption")));
}

#[test]
fn multiple_tables_on_one_slide_round_trip() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    let mut first = Table::new(1, 2).unwrap();
    first.set_cell_text(0, 0, "one").unwrap();
    first.set_cell_text(0, 1, "two").unwrap();
    writer.add_table(slide, 20, 20, first).unwrap();

    let mut second = Table::new(3, 1).unwrap();
    second.set_cell_text(2, 0, "tail").unwrap();
    writer.add_table(slide, 20, 200, second).unwrap();

    let mut package = reopen(write(&mut writer));
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let shapes = slides[0].shapes().unwrap();

    let tables: Vec<_> = shapes.iter().filter_map(|shape| shape.as_table()).collect();
    assert_eq!(tables.len(), 2);
    assert_eq!((tables[0].rows(), tables[0].columns()), (1, 2));
    assert_eq!(tables[0].cell(0, 1), Some("two"));
    assert_eq!((tables[1].rows(), tables[1].columns()), (3, 1));
    assert_eq!(tables[1].cell(2, 0), Some("tail"));
}

#[test]
fn unicode_cell_text_round_trips() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    let mut table = Table::new(1, 1).unwrap();
    table.set_cell_text(0, 0, "表格 Zelle").unwrap();
    writer.add_table(slide, 10, 10, table).unwrap();

    let mut package = reopen(write(&mut writer));
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let shapes = slides[0].shapes().unwrap();
    let read_table = shapes[0].as_table().expect("table group");
    assert_eq!(read_table.cell(0, 0), Some("表格 Zelle"));
}

#[test]
fn custom_cell_dimensions_round_trip_to_bounds() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    let mut table = Table::new(2, 2).unwrap();
    table.set_column_width(0, 120).unwrap();
    table.set_column_width(1, 80).unwrap();
    table.set_row_height(0, 50).unwrap();
    table.set_row_height(1, 30).unwrap();
    table.set_cell_text(0, 0, "anchor").unwrap();
    writer.add_table(slide, 72, 36, table).unwrap();

    let mut package = reopen(write(&mut writer));
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let shapes = slides[0].shapes().unwrap();
    let read_table = shapes[0].as_table().expect("table group");

    // PPT master units: 8 per point. Table at (72pt, 36pt) -> (576, 288);
    // 200pt x 80pt grid -> 1600 x 640 master units.
    assert_eq!((read_table.left(), read_table.top()), (576, 288));
    assert_eq!((read_table.width(), read_table.height()), (1600, 640));
}
