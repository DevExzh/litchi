use std::env;

use litchi_iwa::numbers::{NumbersDocument, NumbersEditor};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    if arguments.len() != 3 {
        return Err(
            "usage: duplicate_numbers_table <input.numbers> <output.numbers> <table-index>".into(),
        );
    }
    let table_index = arguments[2].parse::<usize>()?;
    let mut editor = NumbersEditor::open(&arguments[0])?;
    let source = editor
        .tables()?
        .into_iter()
        .nth(table_index)
        .ok_or("table index was not found")?;
    let created = editor.duplicate_table(source.object_id)?;
    editor.save(&arguments[1])?;

    let document = NumbersDocument::open(&arguments[1])?;
    let sheets = document.sheets()?;
    let copied = sheets
        .iter()
        .flat_map(|sheet| &sheet.tables)
        .find(|table| table.name() == created.name)
        .ok_or("duplicated table was not readable after save")?;
    println!(
        "duplicated table {:?} as {:?}: {} rows x {} columns",
        source.name,
        copied.name(),
        copied.row_count(),
        copied.column_count()
    );
    Ok(())
}
