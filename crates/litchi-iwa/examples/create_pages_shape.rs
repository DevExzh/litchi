//! Create a Pages document and editable rectangle without an input package.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_shape <output.pages> [text]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Built from typed IWA objects".to_owned());
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let body = "Pages rectangle created entirely by litchi-iwa";
    let mut editor = PagesEditor::create_with_text(body)?;
    let created = editor.add_body_rectangle(
        body.encode_utf16().count(),
        &text,
        DrawablePoint { x: 180.0, y: 240.0 },
        DrawableSize {
            width: 300.0,
            height: 150.0,
        },
    )?;
    editor.save(output)?;
    println!(
        "created Pages {:?} {} with storage {} at UTF-16 anchor {}",
        created.kind,
        created.drawable_object_id,
        created.storage.object_id,
        created.anchor_character_index
    );
    Ok(())
}
