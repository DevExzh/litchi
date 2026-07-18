use std::env;
use std::path::PathBuf;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(
        arguments
            .next()
            .ok_or("usage: insert_pages_row <input.pages> <output.pages> <table-id> <row>")?,
    );
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let table_id = arguments.next().ok_or("missing table ID")?.parse::<u64>()?;
    let row = arguments
        .next()
        .ok_or("missing row index")?
        .parse::<usize>()?;
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = PagesEditor::open(input)?;
    editor.insert_table_row(table_id, row)?;
    editor.save(output)?;
    Ok(())
}
