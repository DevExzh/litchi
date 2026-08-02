//! Create Pages, Numbers, and Keynote files with native table-cell fills.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::{CellValue, NumbersDocumentBuilder};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeFill};

const ROW: usize = 1;
const COLUMN: usize = 1;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_fills <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("table-fills.numbers"))?;
    create_pages(&output.join("table-fills.pages"))?;
    create_keynote(&output.join("table-fills.key"))?;
    Ok(())
}

fn fill() -> Result<ShapeFill, Box<dyn std::error::Error>> {
    Ok(ShapeFill::Solid(RgbaColor::new(
        0.96,
        0.72,
        0.12,
        1.0,
        RgbColorSpace::Srgb,
    )?))
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Fills")
        .table_dimensions(3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cell(table_id, ROW, COLUMN, CellValue::Text("Numbers".to_owned()))?;
    editor.set_table_cell_fill(table_id, ROW, COLUMN, &fill()?)?;
    editor.save(output)?;
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Created from scratch with a native table-cell fill.\n")
        .body_table("Fills", 3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(table_id, ROW, COLUMN, CellValue::Text("Pages".to_owned()))?;
    editor.set_table_cell_fill(table_id, ROW, COLUMN, &fill()?)?;
    editor.save(output)?;
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Native table-cell fills")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Fills",
        3,
        3,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        COLUMN,
        CellValue::Text("Keynote".to_owned()),
    )?;
    editor.set_slide_table_cell_fill(0, table.model_object_id, ROW, COLUMN, &fill()?)?;
    editor.save(output)?;
    Ok(())
}
