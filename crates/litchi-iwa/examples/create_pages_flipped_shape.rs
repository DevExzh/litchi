//! Create a Pages shape and apply a native horizontal Arrange flip from scratch.

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawableFlipAxis, DrawablePoint, DrawableSize, ShapePreset};

const ARROW_POSITION: DrawablePoint = DrawablePoint { x: 180.0, y: 240.0 };
const ARROW_SIZE: DrawableSize = DrawableSize {
    width: 300.0,
    height: 150.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_pages_flipped_shape <output.pages>")?;
    let body = "This left-facing arrow was created entirely by litchi-iwa.";
    let mut editor = PagesEditor::create_with_text(body)?;
    let arrow = editor.add_body_shape(
        body.encode_utf16().count(),
        "Horizontally Flipped",
        ARROW_POSITION,
        ARROW_SIZE,
        ShapePreset::RightArrow,
    )?;
    editor.flip_body_shape(arrow.drawable_object_id, DrawableFlipAxis::Horizontal)?;
    editor.save(output)?;
    println!(
        "created horizontally flipped Pages arrow {}",
        arrow.drawable_object_id
    );
    Ok(())
}
