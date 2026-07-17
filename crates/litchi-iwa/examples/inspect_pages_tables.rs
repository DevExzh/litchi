use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: inspect_pages_tables INPUT")?;
    let editor = PagesEditor::open(input)?;
    for table in editor.tables()? {
        println!(
            "anchor={} drawable={} model={} name={:?} dimensions={}x{}",
            table.anchor_character_index,
            table.drawable_object_id,
            table.model_object_id,
            table.name,
            table.rows,
            table.columns,
        );
    }
    Ok(())
}
