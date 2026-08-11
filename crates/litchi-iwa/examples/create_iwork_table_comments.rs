//! Create Pages, Numbers, and Keynote documents with table-comment threads.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

const TABLE_ROWS: usize = 3;
const TABLE_COLUMNS: usize = 3;
const COMMENT_ROW: usize = 1;
const COMMENT_COLUMN: usize = 1;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_comments <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("table-comments.numbers"))?;
    create_pages(&output.join("table-comments.pages"))?;
    create_keynote(&output.join("table-comments.key"))?;
    Ok(())
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Review")
        .table_dimensions(TABLE_ROWS, TABLE_COLUMNS)
        .build()?;
    let table_id = editor.tables()?.remove(0).id();
    editor.set_cell_comment(
        table_id,
        COMMENT_ROW,
        COMMENT_COLUMN,
        "Numbers root comment",
    )?;
    editor.add_cell_comment_reply(
        table_id,
        COMMENT_ROW,
        COMMENT_COLUMN,
        "Numbers direct reply",
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Cross-suite table comments\n")
        .body_table("Review", TABLE_ROWS, TABLE_COLUMNS)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell_comment(table_id, COMMENT_ROW, COMMENT_COLUMN, "Pages root comment")?;
    editor.add_table_cell_comment_reply(
        table_id,
        COMMENT_ROW,
        COMMENT_COLUMN,
        "Pages direct reply",
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Cross-suite table comments")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Review",
        TABLE_ROWS,
        TABLE_COLUMNS,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    editor.set_slide_table_cell_comment(
        0,
        table.model_object_id,
        COMMENT_ROW,
        COMMENT_COLUMN,
        "Keynote root comment",
    )?;
    editor.add_slide_table_cell_comment_reply(
        0,
        table.model_object_id,
        COMMENT_ROW,
        COMMENT_COLUMN,
        "Keynote direct reply",
    )?;
    editor.save(output)?;
    Ok(())
}
