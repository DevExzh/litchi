//! List typed paragraph Drop Caps in every text storage of an iWork package.

use std::env;

use litchi_iwa::text::IWorkTextEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: inspect_iwork_drop_caps <input.pages|input.numbers|input.key>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let editor = IWorkTextEditor::open(input)?;
    for storage in editor.storages()? {
        for placement in editor.paragraph_drop_caps(storage.object_id)? {
            println!(
                "storage={} paragraph_utf16={} {:?}",
                storage.object_id,
                placement.paragraph.utf16_index(),
                placement.drop_cap
            );
        }
    }
    Ok(())
}
