//! Create, update, or delete typed text-language boundaries in any iWork file.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextLanguage, TextPosition};

enum LanguageEdit {
    ResetAll,
    Set(TextPosition, TextLanguage),
    Remove(TextPosition),
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_text_language <input> <output> <storage-id> <reset-all|utf16-position> [automatic|remove|language-tag]",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let target = arguments
        .next()
        .ok_or("missing reset-all or UTF-16 position")?;
    let edit = if target == "reset-all" {
        LanguageEdit::ResetAll
    } else {
        let position = TextPosition::from_utf16_index(target.parse()?)?;
        match arguments.next().as_deref() {
            Some("automatic") => LanguageEdit::Set(position, TextLanguage::Automatic),
            Some("remove") => LanguageEdit::Remove(position),
            Some(tag) => LanguageEdit::Set(position, TextLanguage::tag(tag)?),
            None => return Err("missing language edit".into()),
        }
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    match edit {
        LanguageEdit::ResetAll => {
            editor.reset_text_languages(storage_id)?;
        },
        LanguageEdit::Set(position, language) => {
            editor.set_text_language(storage_id, position, language)?;
        },
        LanguageEdit::Remove(position) => {
            editor.remove_text_language_boundary(storage_id, position)?;
        },
    }
    editor.save(output)?;
    Ok(())
}
