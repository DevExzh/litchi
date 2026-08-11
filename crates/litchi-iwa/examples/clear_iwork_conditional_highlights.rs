//! Clear conditional highlighting from the first table in Pages, Numbers, and Keynote files.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteEditor;
use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa::pages::PagesEditor;

const HIGHLIGHT_ROW: usize = 1;
const HIGHLIGHT_COLUMN: usize = 1;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args().skip(1);
    let source = PathBuf::from(arguments.next().ok_or(
        "usage: clear_iwork_conditional_highlights <source-directory> <output-directory>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or(
        "usage: clear_iwork_conditional_highlights <source-directory> <output-directory>",
    )?);
    std::fs::create_dir_all(&output)?;

    clear_numbers(
        &source.join("conditional-highlight.numbers"),
        &output.join("conditional-highlight.numbers"),
    )?;
    clear_pages(
        &source.join("conditional-highlight.pages"),
        &output.join("conditional-highlight.pages"),
    )?;
    clear_keynote(
        &source.join("conditional-highlight.key"),
        &output.join("conditional-highlight.key"),
    )?;
    Ok(())
}

fn clear_numbers(source: &Path, output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersEditor::open(source)?;
    let table_id = editor.tables()?.remove(0).id();
    editor.clear_cell_conditional_highlighting(table_id, HIGHLIGHT_ROW, HIGHLIGHT_COLUMN)?;
    editor.save(output)?;
    assert!(
        NumbersEditor::open(output)?
            .cell_conditional_highlighting(table_id, HIGHLIGHT_ROW, HIGHLIGHT_COLUMN)?
            .is_none()
    );
    Ok(())
}

fn clear_pages(source: &Path, output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesEditor::open(source)?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.clear_table_cell_conditional_highlighting(table_id, HIGHLIGHT_ROW, HIGHLIGHT_COLUMN)?;
    editor.save(output)?;
    assert!(
        PagesEditor::open(output)?
            .table_cell_conditional_highlighting(table_id, HIGHLIGHT_ROW, HIGHLIGHT_COLUMN)?
            .is_none()
    );
    Ok(())
}

fn clear_keynote(source: &Path, output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteEditor::open(source)?;
    let table_id = editor.slide_tables(0)?.remove(0).model_object_id;
    editor.clear_slide_table_cell_conditional_highlighting(
        0,
        table_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
    )?;
    editor.save(output)?;
    assert!(
        KeynoteEditor::open(output)?
            .slide_table_cell_conditional_highlighting(
                0,
                table_id,
                HIGHLIGHT_ROW,
                HIGHLIGHT_COLUMN,
            )?
            .is_none()
    );
    Ok(())
}
