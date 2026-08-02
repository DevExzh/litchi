//! Remove a slide from an existing Keynote package.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: remove_keynote_slide <input.key> <output.key> <slide-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments.next().ok_or("missing slide index")?.parse()?;

    let mut editor = KeynoteEditor::open(input)?;
    editor.remove_slide(slide_index)?;
    editor.save(output)?;
    Ok(())
}
