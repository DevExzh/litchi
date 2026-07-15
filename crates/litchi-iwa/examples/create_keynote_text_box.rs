//! Create a Keynote presentation and ordinary text box without an input package.

use std::env;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_text_box <output.key> [text]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Built from typed IWA objects".to_owned());
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("No embedded package or source drawable")
        .build()?;
    let created = editor.add_slide_text_box(
        0,
        &text,
        DrawablePoint { x: 144.0, y: 720.0 },
        DrawableSize {
            width: 1_200.0,
            height: 120.0,
        },
    )?;
    editor.save(output)?;
    println!(
        "created Keynote text box {} with storage {}",
        created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}
