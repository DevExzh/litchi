//! Create a Keynote presentation and editable rectangle without an input package.

use std::env;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_shape <output.key> [text]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Built from typed IWA objects".to_owned());
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("Shape built from typed IWA objects")
        .build()?;
    let created = editor.add_slide_rectangle(
        0,
        &text,
        DrawablePoint { x: 720.0, y: 660.0 },
        DrawableSize {
            width: 480.0,
            height: 240.0,
        },
    )?;
    editor.save(output)?;
    println!(
        "created Keynote {:?} {} with storage {}",
        created.kind, created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}
