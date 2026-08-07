use std::env;
use std::path::PathBuf;

use litchi_iwa::numbers::NumbersEditor;
use litchi_numbers::TableSelector;
use litchi_numbers::table::topology::ColumnDeletion;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: remove_numbers_column <input.numbers> <output.numbers> <table-index> <header|body> <index>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let table_index = arguments
        .next()
        .ok_or("missing table index")?
        .parse::<usize>()?;
    let section = arguments.next().ok_or("missing column section")?;
    let index = arguments
        .next()
        .ok_or("missing section-relative column index")?
        .parse::<usize>()?;
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let deletion = match section.as_str() {
        "header" => ColumnDeletion::header(index),
        "body" => ColumnDeletion::body(index),
        _ => return Err(format!("unsupported column section {section:?}").into()),
    };

    let mut editor = NumbersEditor::open(input)?;
    editor.remove_table_column(TableSelector::index(table_index), deletion)?;
    editor.save(output)?;
    Ok(())
}
