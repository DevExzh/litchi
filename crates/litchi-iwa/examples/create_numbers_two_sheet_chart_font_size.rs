//! Generate the tracked two-sheet chart-font-size locality fixture.
//!
//! This host-owned recipe is retained only to make the focused migration test
//! reproducible before W3 removes `litchi-iwa`.

use std::{env, fs};

use litchi_iwa::{
    charts::{ChartData, Kind},
    numbers,
    shapes::{DrawablePoint, DrawableSize},
};
use litchi_numbers::sheet::chart::font::{Property, TextPointSize};

fn set_size(
    source: &[u8],
    sheet: usize,
    points: f32,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_chart_font(sheet, 0usize)?;
    edit.set(Property::Size(TextPointSize::from_points(points)?))?;
    let commit = edit.commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    Ok(bytes)
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = env::args().nth(1).ok_or("usage: generator OUTPUT")?;
    let mut editor = numbers::Builder::new()
        .sheet_name("Primary chart")
        .table_name("Primary data")
        .build()?;
    let first = editor.sheets()?.first().ok_or("missing first sheet")?.id();
    let second = editor.add_empty_sheet("Secondary chart")?;
    let data = ChartData::new(
        vec!["Series".to_owned()],
        vec!["Value".to_owned()],
        vec![vec![Some(42.0)]],
    )?;
    editor.add_sheet_chart(
        first,
        Kind::Column2d,
        data.clone(),
        DrawablePoint { x: 64.0, y: 64.0 },
        DrawableSize {
            width: 320.0,
            height: 220.0,
        },
    )?;
    editor.add_sheet_chart(
        second.id(),
        Kind::Column2d,
        data,
        DrawablePoint { x: 64.0, y: 64.0 },
        DrawableSize {
            width: 320.0,
            height: 220.0,
        },
    )?;

    // Seed distinct values through the focused leaf transaction. The ops test
    // can therefore prove that a retained first-sheet node is not retargeted
    // when the second-sheet cache entry is initialized later.
    let source = editor.to_bytes()?;
    let source = set_size(&source, 0, 13.0)?;
    let source = set_size(&source, 1, 27.0)?;
    fs::write(output, source)?;
    Ok(())
}
