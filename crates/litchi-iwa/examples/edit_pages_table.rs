use std::env;

use litchi_iwa::pages::{PagesCellValue, PagesEditor};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let input = args.next().ok_or("usage: edit_pages_table INPUT OUTPUT")?;
    let output = args.next().ok_or("usage: edit_pages_table INPUT OUTPUT")?;
    let mut editor = PagesEditor::open(input)?;
    let table = editor
        .tables()?
        .into_iter()
        .next()
        .ok_or("the Pages body contains no table")?;
    editor.set_table_cell(
        table.model_object_id,
        2,
        0,
        PagesCellValue::Text("Updated by litchi-iwa".to_owned()),
    )?;
    editor.rename_table(table.model_object_id, "Edited Table")?;
    editor.resize_table(
        table.model_object_id,
        table.rows.max(5),
        table.columns.max(4),
    )?;
    editor.save(output)?;
    Ok(())
}
