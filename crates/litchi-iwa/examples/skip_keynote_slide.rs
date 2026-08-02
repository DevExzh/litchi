//! Set a Keynote slide's skipped state.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: skip_keynote_slide <input.key> <output.key> <slide-index> <skipped>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments
        .next()
        .ok_or("missing slide index")?
        .parse::<usize>()?;
    let skipped = arguments
        .next()
        .ok_or("missing skipped flag")?
        .parse::<bool>()?;

    let mut editor = KeynoteEditor::open(input)?;
    editor.set_slide_skipped(slide_index, skipped)?;
    editor.save(output)?;
    Ok(())
}
