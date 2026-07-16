//! Create a Keynote presentation and ordinary text box without an input package.

use std::env;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, ShapeTextAutoSize, ShapeTextInset, ShapeTextInsets,
    ShapeTextLayout, ShapeTextVerticalAlignment,
};
use litchi_iwa::text::{TextAlignment, TextColumnCount, TextColumns};

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
    editor.set_slide_text_box_columns(
        0,
        created.drawable_object_id,
        &TextColumns::equal(TextColumnCount::new(4)?, None),
    )?;
    editor.set_slide_text_box_text_layout(
        0,
        created.drawable_object_id,
        ShapeTextLayout::new(
            ShapeTextVerticalAlignment::Middle,
            ShapeTextInsets::uniform(ShapeTextInset::from_points(12.0)?),
            ShapeTextAutoSize::ShrinkToFit,
        ),
    )?;
    editor.set_slide_text_box_paragraph_alignment(
        0,
        created.drawable_object_id,
        TextAlignment::Justified,
    )?;
    editor.save(output)?;
    println!(
        "created four-column Keynote text box {} with storage {}",
        created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}
