//! Create editable Pages, Numbers, and Keynote files without input documents.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_iwa::numbers::{CellValue, NumbersEditor};
use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output_directory = env::args()
        .nth(1)
        .ok_or("usage: create_iwork <output-directory>")?;
    let output_directory = std::path::Path::new(&output_directory);
    std::fs::create_dir_all(output_directory)?;

    let mut pages = PagesEditor::create()?;
    pages.set_body_text("Created from scratch with litchi-iwa")?;
    pages.save(output_directory.join("created.pages"))?;

    let mut numbers = NumbersEditor::create()?;
    let table_id = numbers.tables()?.remove(0).object_id;
    numbers.set_cell(table_id, 0, 0, CellValue::Text("Created".to_owned()))?;
    numbers.set_cell(table_id, 0, 1, CellValue::Number(42.0))?;
    numbers.save(output_directory.join("created.numbers"))?;

    let mut keynote = KeynoteEditor::create()?;
    keynote.set_slide_title(0, "Created from scratch")?;
    keynote.set_slide_body(0, "litchi-iwa")?;
    keynote.save(output_directory.join("created.key"))?;
    Ok(())
}
