//! Create Pages, Numbers, and Keynote files with explicit native cell borders.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::numbers::editor::table::cell::BorderSide;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeStroke, StrokePattern, StrokeWidth,
};
use litchi_numbers::cell::Value as CellValue;
const ROW: usize = 1;
const COLUMN: usize = 1;
const SIDES: [BorderSide; 4] = [
    BorderSide::Left,
    BorderSide::Right,
    BorderSide::Top,
    BorderSide::Bottom,
];

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_borders <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("table-borders.numbers"))?;
    create_pages(&output.join("table-borders.pages"))?;
    create_keynote(&output.join("table-borders.key"))?;
    Ok(())
}

fn border() -> Result<ShapeStroke, Box<dyn std::error::Error>> {
    Ok(ShapeStroke::new(
        RgbaColor::new(0.86, 0.12, 0.18, 1.0, RgbColorSpace::Srgb)?,
        StrokeWidth::new(4.0)?,
        StrokePattern::Solid,
    ))
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Borders")
        .table_dimensions(3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cell(table_id, ROW, COLUMN, CellValue::Text("Numbers".to_owned()))?;
    for side in SIDES {
        editor.set_table_cell_border(table_id, ROW, COLUMN, side, border()?)?;
    }
    editor.save(output)?;
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Created from scratch with explicit native borders.\n")
        .body_table("Borders", 3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(table_id, ROW, COLUMN, CellValue::Text("Pages".to_owned()))?;
    for side in SIDES {
        editor.set_table_cell_border(table_id, ROW, COLUMN, side, border()?)?;
    }
    editor.save(output)?;
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Explicit native cell borders")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Borders",
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
    for side in SIDES {
        editor.set_slide_table_cell_border(
            0,
            table.model_object_id,
            ROW,
            COLUMN,
            side,
            border()?,
        )?;
    }
    editor.save(output)?;
    Ok(())
}
