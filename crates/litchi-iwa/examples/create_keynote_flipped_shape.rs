//! Create a Keynote shape and apply a native horizontal Arrange flip from scratch.

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawableFlipAxis, DrawablePoint, DrawableSize};
use litchi_iwa_common::shape::path::Preset;

const ARROW_POSITION: DrawablePoint = DrawablePoint { x: 720.0, y: 660.0 };
const ARROW_SIZE: DrawableSize = DrawableSize {
    width: 480.0,
    height: 240.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_keynote_flipped_shape <output.key>")?;
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Flipped Shape")
        .subtitle("Typed native Arrange flip")
        .build()?;
    let arrow = editor.add_slide_shape(
        0,
        "Horizontally Flipped",
        ARROW_POSITION,
        ARROW_SIZE,
        Preset::RightArrow,
    )?;
    editor.flip_slide_shape(0, arrow.drawable_object_id, DrawableFlipAxis::Horizontal)?;
    editor.save(output)?;
    println!(
        "created horizontally flipped Keynote arrow {}",
        arrow.drawable_object_id
    );
    Ok(())
}
