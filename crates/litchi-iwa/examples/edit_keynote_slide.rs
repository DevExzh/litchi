//! Replace a slide's title and body text in an existing Keynote package.

use std::env;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: edit_keynote_slide <input.key> <output.key> <slide-index> <title> <body>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let slide_index = arguments.next().ok_or("missing slide index")?.parse()?;
    let title = arguments.next().ok_or("missing title")?;
    let body = arguments.next().ok_or("missing body")?;

    let mut editor = KeynoteEditor::open(input)?;
    editor.set_slide_title(slide_index, &title)?;
    editor.set_slide_body(slide_index, &body)?;
    editor.save(output)?;
    Ok(())
}
