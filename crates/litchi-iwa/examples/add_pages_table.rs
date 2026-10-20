use std::env;

use litchi_iwa::pages::{self};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let input = args.next().ok_or("usage: add_pages_table INPUT OUTPUT")?;
    let output = args.next().ok_or("usage: add_pages_table INPUT OUTPUT")?;
    let mut editor = pages::Editor::open(input)?;
    let anchor = editor.body_text()?.encode_utf16().count();
    let table = editor.add_table(anchor, "Added by litchi-iwa", 3, 2)?;
    editor.set_table_cell(
        table.model_object_id,
        0,
        0,
        pages::CellValue::Text("Independent native storage".to_owned()),
    )?;
    editor.save(output)?;
    Ok(())
}
