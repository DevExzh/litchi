//! Set or reset one paragraph's typed list nesting level.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, ParagraphListLevel, ParagraphStart};

enum LevelEdit {
    Reset,
    Set(ParagraphListLevel),
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_paragraph_list_level <input> <output> <storage-id> <utf16-paragraph-start> <reset|0..8>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let paragraph = ParagraphStart::from_utf16_index(
        arguments
            .next()
            .ok_or("missing UTF-16 paragraph start")?
            .parse()?,
    )?;
    let edit = match arguments.next().as_deref() {
        Some("reset") => LevelEdit::Reset,
        Some(level) => LevelEdit::Set(ParagraphListLevel::new(level.parse()?)?),
        None => return Err("missing list level".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    match edit {
        LevelEdit::Reset => {
            editor.reset_paragraph_list_level(storage_id, paragraph)?;
        },
        LevelEdit::Set(level) => {
            editor.set_paragraph_list_level(storage_id, paragraph, level)?;
        },
    }
    editor.save(output)?;
    Ok(())
}
