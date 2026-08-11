//! Remove an ordinary Numbers shape and its private object graph.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: remove_numbers_shape <input.numbers> <output.numbers> <sheet-index> <shape-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let shape_index: usize = arguments.next().ok_or("missing shape index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet_id = editor
        .sheets()?
        .get(sheet_index)
        .ok_or("sheet index is out of bounds")?
        .id();
    let shape = editor
        .sheet_shapes(sheet_id)?
        .get(shape_index)
        .cloned()
        .ok_or("shape index is out of bounds")?;
    let removed = editor.remove_sheet_shape(sheet_id, shape.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "sheet={sheet_id} drawable={} storage={} kind={:?}",
        removed.shape.drawable_object_id, removed.shape.storage.id, removed.shape.kind
    );
    Ok(())
}
