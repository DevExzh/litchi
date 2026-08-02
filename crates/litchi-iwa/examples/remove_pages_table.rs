use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let input = args
        .next()
        .ok_or("usage: remove_pages_table INPUT OUTPUT")?;
    let output = args
        .next()
        .ok_or("usage: remove_pages_table INPUT OUTPUT")?;
    let mut editor = PagesEditor::open(input)?;
    let table = editor
        .tables()?
        .into_iter()
        .next()
        .ok_or("the Pages body contains no table")?;
    editor.remove_table(table.model_object_id)?;
    editor.save(output)?;
    Ok(())
}
