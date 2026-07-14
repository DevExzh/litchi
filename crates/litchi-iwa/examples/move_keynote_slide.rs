//! Move a slide in an existing Keynote package.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: move_keynote_slide <input.key> <output.key> <from-index> <to-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let from = arguments.next().ok_or("missing source index")?.parse()?;
    let to = arguments
        .next()
        .ok_or("missing destination index")?
        .parse()?;

    let mut editor = KeynoteEditor::open(input)?;
    editor.move_slide(from, to)?;
    editor.save(output)?;
    Ok(())
}
