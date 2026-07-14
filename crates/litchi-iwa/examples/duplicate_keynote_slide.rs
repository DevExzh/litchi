//! Duplicate a slide in an existing Keynote package.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: duplicate_keynote_slide <input.key> <output.key> <slide-index>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments.next().ok_or("missing slide index")?.parse()?;

    let mut editor = KeynoteEditor::open(input)?;
    let slide = editor.duplicate_slide(slide_index)?;
    editor.save(output)?;
    println!(
        "duplicated slide {} as slide {} (node {}, object {})",
        slide_index, slide.index, slide.node_id, slide.slide_id
    );
    Ok(())
}
