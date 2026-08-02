//! Set one uniform typed font identity in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextFont};

enum FontEdit {
    Reset,
    Set(TextFont),
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_text_font <input> <output> <storage-id> <inherit|default|font-name>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let edit = match arguments.next().as_deref() {
        Some("inherit") => FontEdit::Reset,
        Some("default") => FontEdit::Set(TextFont::Default),
        Some(name) => FontEdit::Set(TextFont::named(name)?),
        None => return Err("missing font".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    match edit {
        FontEdit::Reset => {
            editor.reset_text_font(storage_id)?;
        },
        FontEdit::Set(font) => editor.set_text_font(storage_id, font)?,
    }
    editor.save(output)?;
    Ok(())
}
