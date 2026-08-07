//! Set one uniform typed outline in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, Outline};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: edit_iwork_text_outline <input> <output> <storage-id> <none|standard>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let outline = match arguments.next().as_deref() {
        Some("none") => Outline::None,
        Some("standard") => Outline::standard(),
        Some(_) => return Err("outline must be none or standard".into()),
        None => return Err("missing outline".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_outline(storage_id, outline)?;
    editor.save(output)?;
    Ok(())
}
