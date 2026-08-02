//! Append an empty sheet to an existing Numbers workbook.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: add_numbers_sheet <input.numbers> <output.numbers> <sheet-name>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let name = arguments.next().ok_or("missing sheet name")?;

    let mut editor = NumbersEditor::open(input)?;
    editor.add_empty_sheet(&name)?;
    editor.save(output)?;
    Ok(())
}
