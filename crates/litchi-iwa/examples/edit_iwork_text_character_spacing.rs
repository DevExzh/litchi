//! Set uniform typed character spacing in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextCharacterSpacing};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_text_character_spacing <input> <output> <storage-id> <percent>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let spacing = TextCharacterSpacing::from_percent(
        arguments
            .next()
            .ok_or("missing character spacing percentage")?
            .parse()?,
    )?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_character_spacing(storage_id, spacing)?;
    editor.save(output)?;
    Ok(())
}
