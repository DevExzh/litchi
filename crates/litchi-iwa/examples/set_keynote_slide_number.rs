//! Show or hide the layout-provided slide number on one Keynote slide.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: set_keynote_slide_number <input.key> <output.key> <slide-index> <true|false>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments.next().ok_or("missing slide index")?.parse()?;
    let visible = arguments
        .next()
        .ok_or("missing slide-number visibility")?
        .parse()?;

    let mut editor = KeynoteEditor::open(input)?;
    editor.set_slide_number_visible(slide_index, visible)?;
    editor.save(output)?;
    Ok(())
}
