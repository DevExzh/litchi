#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "table edit tests use panic-driven assertions"
)]

//! Table-cell text through the canonical primary edit (record 0323).

use litchi_ppt::Package;
use litchi_ppt::package::edit::{Position, Target};
use std::path::PathBuf;

fn fixture(path: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    std::fs::read(path).expect("PPT fixture")
}

const TABLES: &str = "poi/test-data/slideshow/bug61881.ppt";

fn table(package: &Package, slide: usize) -> (usize, Vec<Vec<String>>, Vec<i32>, Vec<i32>) {
    let presentation = package.presentation().unwrap();
    let slide = presentation.slide_at(slide).unwrap().unwrap();
    let shapes = slide.shapes().unwrap();
    let (index, table) = shapes
        .iter()
        .enumerate()
        .find_map(|(index, shape)| shape.as_table().map(|table| (index, table)))
        .expect("table");
    let cells = (0..table.rows())
        .map(|row| {
            (0..table.columns())
                .map(|column| table.cell(row, column).unwrap().to_owned())
                .collect()
        })
        .collect();
    (index, cells, table.row_heights(), table.column_widths())
}

fn built() -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    writer.add_textbox(slide, 10, 10, 240, 40, "Title").unwrap();
    let mut table = litchi_ppt::writer::Table::new(2, 3).unwrap();
    for (row, texts) in [["A1", "B1", "C1"], ["A2", "B2", "C2"]].iter().enumerate() {
        for (column, text) in texts.iter().enumerate() {
            table.set_cell_text(row, column, *text).unwrap();
        }
    }
    writer.add_table(slide, 72, 120, table).unwrap();
    let mut output = std::io::Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
    output.into_inner()
}

#[test]
fn cell_text_changes_length_and_keeps_every_other_cell() {
    let source_bytes = built();
    let source = Package::from_bytes(&source_bytes).unwrap();
    let (index, before, heights, widths) = table(&source, 0);
    assert_eq!(heights.len(), 2);
    assert_eq!(widths.len(), 3);
    let mut edit = source.edit().unwrap();
    edit.set_table_cell_text(
        Target::new(Position::new(0), Position::new(index)),
        Position::new(1),
        Position::new(2),
        "A much longer cell",
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let published = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    let (_, after, after_heights, after_widths) = table(&published, 0);
    let mut expected = before.clone();
    expected[1][2] = "A much longer cell".into();
    assert_eq!(after, expected);
    assert_eq!((after_heights, after_widths), (heights, widths));
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source_bytes
    );

    // Equal text is an exact no-op; an absent row is refused.
    let mut edit = source.edit().unwrap();
    edit.set_table_cell_text(
        Target::new(Position::new(0), Position::new(index)),
        Position::new(0),
        Position::new(0),
        "A1",
    )
    .unwrap();
    assert!(edit.commit().unwrap().patch().is_noop());
    let mut edit = source.edit().unwrap();
    assert!(
        edit.set_table_cell_text(
            Target::new(Position::new(0), Position::new(index)),
            Position::new(2),
            Position::new(0),
            "x",
        )
        .is_err()
    );
}

#[test]
fn empty_powerpoint_cells_gain_a_text_atom() {
    let source_bytes = fixture(TABLES);
    let source = Package::from_bytes(&source_bytes).unwrap();
    let (index, before, _, _) = table(&source, 10);
    assert_eq!(before.len(), 5);
    let slide_text = |package: &Package, slide: usize| {
        package
            .presentation()
            .unwrap()
            .slide_at(slide)
            .unwrap()
            .unwrap()
            .text()
            .unwrap()
            .to_owned()
    };
    let mut edit = source.edit().unwrap();
    for (row, text) in ["one", "two"].into_iter().enumerate() {
        edit.set_table_cell_text(
            Target::new(Position::new(10), Position::new(index)),
            Position::new(row),
            Position::new(0),
            text,
        )
        .unwrap();
    }
    let commit = edit.commit().unwrap();
    let published = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    let (_, after, _, _) = table(&published, 10);
    assert_eq!(after[0][0], "one");
    assert_eq!(after[1][0], "two");
    assert_eq!(after[2..], before[2..]);
    assert_eq!(slide_text(&published, 9), slide_text(&source, 9));
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
}
