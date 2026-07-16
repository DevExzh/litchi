//! Set one uniform typed ligature policy in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextLigatures};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_text_ligatures <input> <output> <storage-id> <required|standard|all>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let ligatures = match arguments.next().as_deref() {
        Some("required") => TextLigatures::RequiredOnly,
        Some("standard") => TextLigatures::Standard,
        Some("all") => TextLigatures::All,
        Some(_) => return Err("ligatures must be required, standard, or all".into()),
        None => return Err("missing ligature policy".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_ligatures(storage_id, ligatures)?;
    editor.save(output)?;
    Ok(())
}
