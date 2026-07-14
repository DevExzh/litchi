use std::env;

use litchi_iwa::numbers::{NumbersDocument, NumbersEditor};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    if arguments.len() != 3 {
        return Err(
            "usage: duplicate_numbers_sheet <input.numbers> <output.numbers> <sheet-index>".into(),
        );
    }
    let sheet_index = arguments[2].parse::<usize>()?;
    let mut editor = NumbersEditor::open(&arguments[0])?;
    let source = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index was not found")?;
    let created = editor.duplicate_sheet(source.object_id)?;
    editor.save(&arguments[1])?;

    let document = NumbersDocument::open(&arguments[1])?;
    let sheets = document.sheets()?;
    let copied = sheets
        .get(created.index)
        .filter(|sheet| sheet.name == created.name)
        .ok_or("duplicated sheet was not readable after save")?;
    println!(
        "duplicated sheet {:?} as {:?} with {} tables",
        source.name,
        copied.name,
        copied.tables.len()
    );
    Ok(())
}
