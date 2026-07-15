use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: add_pages_text_box <input.pages> <output.pages> <text>")?;
    let output = arguments
        .next()
        .ok_or("usage: add_pages_text_box <input.pages> <output.pages> <text>")?;
    let text = arguments
        .next()
        .ok_or("usage: add_pages_text_box <input.pages> <output.pages> <text>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let anchor = editor.body_text()?.encode_utf16().count();
    let created = editor.add_text_box(
        anchor,
        &text,
        DrawablePoint { x: 96.0, y: 144.0 },
        DrawableSize {
            width: 240.0,
            height: 72.0,
        },
    )?;
    editor.save(output)?;
    println!(
        "created Pages text box {} with storage {}",
        created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}
