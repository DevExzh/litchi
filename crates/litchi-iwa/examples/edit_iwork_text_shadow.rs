//! Set one uniform typed drop shadow in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, Shadow};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: edit_iwork_text_shadow <input> <output> <storage-id> <none|standard>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let shadow = match arguments.next().as_deref() {
        Some("none") => Shadow::None,
        Some("standard") => Shadow::standard(),
        Some(_) => return Err("shadow must be none or standard".into()),
        None => return Err("missing shadow".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_shadow(storage_id, shadow)?;
    editor.save(output)?;
    Ok(())
}
