//! Move a Numbers sheet within workbook order.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: move_numbers_sheet <input.numbers> <output.numbers> <from> <to>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let from: usize = arguments.next().ok_or("missing source index")?.parse()?;
    let to: usize = arguments
        .next()
        .ok_or("missing destination index")?
        .parse()?;

    let mut editor = NumbersEditor::open(input)?;
    editor.move_sheet(from, to)?;
    editor.save(output)?;
    Ok(())
}
