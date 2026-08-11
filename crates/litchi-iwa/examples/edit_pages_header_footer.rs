//! Replace one reachable Pages header/footer text storage.

use std::env;

use litchi_iwa::{pages::PagesEditor, text::TextStorageId};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_header_footer <input.pages> <output.pages> <storage-id> <text>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments
        .next()
        .ok_or("missing storage ID")?
        .parse::<TextStorageId>()?;
    let replacement = arguments.next().ok_or("missing replacement text")?;

    let mut editor = PagesEditor::open(input)?;
    editor.set_header_footer_text(storage_id, &replacement)?;
    editor.save(output)?;
    Ok(())
}
