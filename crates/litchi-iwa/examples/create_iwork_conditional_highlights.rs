//! Create Pages, Numbers, and Keynote files with native conditional highlighting.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::{CellValue, NumbersDocumentBuilder};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
use litchi_iwa::table_cell_conditional_highlight::{
    TableCellConditionalHighlightCondition, TableCellConditionalHighlightNumber,
    TableCellConditionalHighlightRule, TableCellConditionalHighlightStyle,
};

const HIGHLIGHT_ROW: usize = 1;
const HIGHLIGHT_COLUMN: usize = 1;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_conditional_highlights <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("conditional-highlight.numbers"))?;
    create_pages(&output.join("conditional-highlight.pages"))?;
    create_keynote(&output.join("conditional-highlight.key"))?;
    Ok(())
}

fn highlight_rule() -> Result<TableCellConditionalHighlightRule, Box<dyn std::error::Error>> {
    let zero = TableCellConditionalHighlightNumber::new(0.0)?;
    let red = RgbaColor::new(0.96, 0.22, 0.18, 1.0, RgbColorSpace::Srgb)?;
    Ok(TableCellConditionalHighlightRule::new(
        TableCellConditionalHighlightCondition::LessThan(zero),
        TableCellConditionalHighlightStyle::new(Some(red), None, true)?,
    ))
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Conditional")
        .table_dimensions(3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cell(
        table_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        CellValue::Number(-5.0),
    )?;
    editor.set_cell_conditional_highlighting(
        table_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        &[highlight_rule()?],
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Conditional highlighting created from scratch.\n")
        .body_table("Conditional", 3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(
        table_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        CellValue::Number(-5.0),
    )?;
    editor.set_table_cell_conditional_highlighting(
        table_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        &[highlight_rule()?],
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Conditional highlighting")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Conditional",
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
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        CellValue::Number(-5.0),
    )?;
    editor.set_slide_table_cell_conditional_highlighting(
        0,
        table.model_object_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        &[highlight_rule()?],
    )?;
    editor.save(output)?;
    Ok(())
}
