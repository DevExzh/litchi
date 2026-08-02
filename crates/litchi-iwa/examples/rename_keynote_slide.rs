//! Set or clear a slide's navigator name in an existing Keynote package.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: rename_keynote_slide <input.key> <output.key> <slide-index> [name]")?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments.next().ok_or("missing slide index")?.parse()?;
    let name = arguments.next();

    let mut editor = KeynoteEditor::open(input)?;
    editor.set_slide_name(slide_index, name.as_deref())?;
    editor.save(output)?;
    Ok(())
}
