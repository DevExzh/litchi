use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = env::args()
        .nth(1)
        .ok_or("usage: list_keynote_tables <presentation.key>")?;
    let editor = KeynoteEditor::open(input)?;
    for slide in editor.slides()? {
        for info in editor.slide_tables(slide.index)? {
            let table = editor.slide_table(slide.index, info.model_object_id)?;
            println!(
                "slide={} drawable={} model={} name={:?} rows={} columns={} cells={:?}",
                slide.index + 1,
                info.drawable_object_id,
                info.model_object_id,
                info.name,
                info.rows,
                info.columns,
                table.cells
            );
        }
    }
    Ok(())
}
