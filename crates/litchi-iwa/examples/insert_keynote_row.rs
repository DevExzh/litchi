use std::env;
use std::path::PathBuf;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: insert_keynote_row <input.key> <output.key> <slide-index> <table-id> <row>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let slide_index = arguments
        .next()
        .ok_or("missing slide index")?
        .parse::<usize>()?;
    let table_id = arguments.next().ok_or("missing table ID")?.parse::<u64>()?;
    let row = arguments
        .next()
        .ok_or("missing row index")?
        .parse::<usize>()?;
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = KeynoteEditor::open(input)?;
    editor.insert_slide_table_row(slide_index, table_id, row)?;
    editor.save(output)?;
    Ok(())
}
