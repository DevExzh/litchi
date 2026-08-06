//! Duplicate an ordinary Numbers text box with independent storage.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: duplicate_numbers_text_box <input.numbers> <output.numbers> <sheet-index> <text-box-index> <text>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let text_box_index: usize = arguments.next().ok_or("missing text-box index")?.parse()?;
    let text = arguments.next().ok_or("missing clone text")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index out of range")?;
    let source = editor
        .sheet_text_boxes(sheet.object_id)?
        .into_iter()
        .nth(text_box_index)
        .ok_or("text-box index out of range")?;
    let created =
        editor.duplicate_sheet_text_box(sheet.object_id, source.drawable_object_id, &text)?;
    editor.save(output)?;
    println!(
        "sheet={sheet_index} source={} clone={} storage={} text={:?}",
        source.drawable_object_id,
        created.drawable_object_id,
        created.storage.id,
        created.storage.storage.text()
    );
    Ok(())
}
