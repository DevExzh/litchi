//! Remove a non-final sheet from an existing Numbers workbook.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: remove_numbers_sheet <input.numbers> <output.numbers> <sheet-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .get(sheet_index)
        .cloned()
        .ok_or("sheet index is out of bounds")?;
    editor.remove_sheet(sheet.object_id)?;
    editor.save(output)?;
    Ok(())
}
