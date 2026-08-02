//! Set one uniform typed baseline script in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextScript};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: edit_iwork_text_script <input> <output> <storage-id> <normal|super|sub>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let script = match arguments.next().as_deref() {
        Some("normal") => TextScript::Normal,
        Some("super") => TextScript::Superscript,
        Some("sub") => TextScript::Subscript,
        Some(_) => return Err("script must be normal, super, or sub".into()),
        None => return Err("missing script mode".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_script(storage_id, script)?;
    editor.save(output)?;
    Ok(())
}
