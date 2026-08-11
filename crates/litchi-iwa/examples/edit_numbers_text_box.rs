//! Replace all text in an ordinary Numbers text box.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_text_box <input.numbers> <output.numbers> <sheet-index> <text-box-index> <replacement>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let text_box_index: usize = arguments.next().ok_or("missing text-box index")?.parse()?;
    let replacement = arguments.next().ok_or("missing replacement text")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index out of range")?;
    let text_box = editor
        .sheet_text_boxes(sheet.id())?
        .into_iter()
        .nth(text_box_index)
        .ok_or("text-box index out of range")?;
    editor.set_sheet_text_box_text(sheet.id(), text_box.drawable_object_id, &replacement)?;
    editor.save(output)?;
    println!(
        "sheet={sheet_index} drawable={} storage={} text={replacement:?}",
        text_box.drawable_object_id, text_box.storage.id
    );
    Ok(())
}
