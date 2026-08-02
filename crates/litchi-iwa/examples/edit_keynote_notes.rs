//! Replace one slide's speaker notes in an existing Keynote package.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: edit_keynote_notes <input.key> <output.key> <slide-index> <notes>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments.next().ok_or("missing slide index")?.parse()?;
    let notes = arguments.next().ok_or("missing notes")?;

    let mut editor = KeynoteEditor::open(input)?;
    editor.set_slide_notes(slide_index, &notes)?;
    editor.save(output)?;
    Ok(())
}
