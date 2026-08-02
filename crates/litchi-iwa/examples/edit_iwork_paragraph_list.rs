//! Apply or remove one canonical list preset in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, ParagraphList};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_paragraph_list <input> <output> <storage-id> <none|bullet|numbered>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let list = match arguments.next().as_deref() {
        Some("none") => ParagraphList::None,
        Some("bullet") => ParagraphList::Bullet,
        Some("numbered") => ParagraphList::Numbered,
        Some(value) => return Err(format!("unsupported list preset {value:?}").into()),
        None => return Err("missing list preset".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_paragraph_list(storage_id, list)?;
    editor.save(output)?;
    Ok(())
}
