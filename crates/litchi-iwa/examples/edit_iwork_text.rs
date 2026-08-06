//! List or replace shared iWork text-storage content.

use std::env;

use litchi_iwa::text::IWorkTextEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: edit_iwork_text <input> [<output> <object-id> <text>]")?;
    let mut editor = IWorkTextEditor::open(&input)?;
    let Some(output) = arguments.next() else {
        for storage in editor.storages()? {
            if !storage.storage.is_empty() {
                println!("{}\t{:?}", storage.object_id, storage.storage.text());
            }
        }
        return Ok(());
    };
    let object_id = arguments.next().ok_or("missing object ID")?.parse()?;
    let text = arguments.next().ok_or("missing replacement text")?;
    editor.set_text(object_id, &text)?;
    editor.save(output)?;
    Ok(())
}
