use std::env;
use std::path::PathBuf;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: remove_numbers_column <input.numbers> <output.numbers> <table-id> <column>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let table_id = arguments.next().ok_or("missing table ID")?.parse::<u64>()?;
    let column = arguments
        .next()
        .ok_or("missing column index")?
        .parse::<usize>()?;
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    editor.remove_table_column(table_id, column)?;
    editor.save(output)?;
    Ok(())
}
