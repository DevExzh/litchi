//! Set one uniform typed capitalization mode in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextCapitalization};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_text_capitalization <input> <output> <storage-id> <none|all|small|title|start>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let capitalization = match arguments.next().as_deref() {
        Some("none") => TextCapitalization::None,
        Some("all") => TextCapitalization::AllCaps,
        Some("small") => TextCapitalization::SmallCaps,
        Some("title") => TextCapitalization::TitleCase,
        Some("start") => TextCapitalization::StartCase,
        Some(_) => return Err("capitalization must be none, all, small, title, or start".into()),
        None => return Err("missing capitalization mode".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_capitalization(storage_id, capitalization)?;
    editor.save(output)?;
    Ok(())
}
