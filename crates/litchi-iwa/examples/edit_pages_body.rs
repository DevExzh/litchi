//! Replace the main body text in an existing Pages package.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: edit_pages_body <input.pages> <output.pages> <text>")?;
    let output = arguments.next().ok_or("missing output path")?;
    let replacement = arguments.next().ok_or("missing replacement text")?;

    let mut editor = PagesEditor::open(input)?;
    editor.set_body_text(&replacement)?;
    editor.save(output)?;
    Ok(())
}
