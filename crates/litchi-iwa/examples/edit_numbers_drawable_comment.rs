//! Create, replace, or clear a sheet-owned Numbers drawable comment.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_drawable_comment <input.numbers> <output.numbers> <sheet-index> <drawable-index> <text|--clear>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let drawable_index: usize = arguments.next().ok_or("missing drawable index")?.parse()?;
    let replacement = arguments
        .next()
        .ok_or("missing replacement text or --clear")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index out of range")?;
    let drawable = editor
        .sheet_drawables(sheet.object_id)?
        .into_iter()
        .nth(drawable_index)
        .ok_or("drawable index out of range")?;
    let drawable_id = drawable.id.get();
    let old = editor.sheet_drawable_comment(sheet.object_id, drawable_id)?;
    if replacement == "--clear" {
        editor.clear_sheet_drawable_comment(sheet.object_id, drawable_id)?;
    } else {
        editor.set_sheet_drawable_comment(sheet.object_id, drawable_id, replacement)?;
    }
    editor.save(output)?;
    let new = editor.sheet_drawable_comment(sheet.object_id, drawable_id)?;
    println!(
        "sheet={sheet_index} drawable={} old={:?} new={:?}",
        drawable_id,
        old.as_ref().map(|value| value.comment.text.as_str()),
        new.as_ref().map(|value| value.comment.text.as_str())
    );
    Ok(())
}
