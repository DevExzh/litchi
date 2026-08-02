//! Remove an ordinary Numbers text box and its private graph.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: remove_numbers_text_box <input.numbers> <output.numbers> <sheet-index> <text-box-index>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let text_box_index: usize = arguments.next().ok_or("missing text-box index")?.parse()?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index out of range")?;
    let target = editor
        .sheet_text_boxes(sheet.object_id)?
        .into_iter()
        .nth(text_box_index)
        .ok_or("text-box index out of range")?;
    let removed = editor.remove_sheet_text_box(sheet.object_id, target.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "sheet={sheet_index} removed_drawable={} removed_storage={} text={:?}",
        removed.text_box.drawable_object_id,
        removed.text_box.storage.object_id,
        removed.text_box.storage.text
    );
    Ok(())
}
