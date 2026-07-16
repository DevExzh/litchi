//! Create a Pages document and multi-column text box without an input package.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, ShapeTextAutoSize, ShapeTextInset, ShapeTextInsets,
    ShapeTextLayout, ShapeTextVerticalAlignment,
};
use litchi_iwa::text::{TextColumnCount, TextColumnGap, TextColumns};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_text_box <output.pages> [text]")?;
    let text = arguments.next().unwrap_or_else(|| {
        "A typed, two-column Pages text box created entirely from scratch. ".repeat(12)
    });
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::create_with_text("Multi-column text box")?;
    let anchor = editor.body_text()?.encode_utf16().count();
    let created = editor.add_text_box(
        anchor,
        &text,
        DrawablePoint { x: 72.0, y: 144.0 },
        DrawableSize {
            width: 468.0,
            height: 360.0,
        },
    )?;
    editor.set_text_box_columns(
        created.drawable_object_id,
        &TextColumns::equal(
            TextColumnCount::new(2)?,
            Some(TextColumnGap::from_points(18.0)?),
        ),
    )?;
    editor.set_text_box_text_layout(
        created.drawable_object_id,
        ShapeTextLayout::new(
            ShapeTextVerticalAlignment::Middle,
            ShapeTextInsets::uniform(ShapeTextInset::from_points(9.0)?),
            ShapeTextAutoSize::ShrinkToFit,
        ),
    )?;
    editor.save(output)?;
    println!(
        "created two-column Pages text box {} with storage {}",
        created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}
