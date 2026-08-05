//! Create overlapping Pages shapes and move one to the front without a template.

use litchi_iwa::DrawableLayerMove;
use litchi_iwa::comments::DrawableObjectId;
use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeFill};
use litchi_iwa_common::shape::path::Preset;

const OVERLAP_POSITION: DrawablePoint = DrawablePoint { x: 180.0, y: 240.0 };
const OVERLAP_SIZE: DrawableSize = DrawableSize {
    width: 300.0,
    height: 150.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_pages_stacked_shapes <output.pages>")?;
    let body = "Layered Pages shapes created entirely by litchi-iwa";
    let mut editor = PagesEditor::create_with_text(body)?;
    let first = editor.add_body_shape_with_fill(
        body.encode_utf16().count(),
        "Moved to Front",
        OVERLAP_POSITION,
        OVERLAP_SIZE,
        Preset::Rectangle,
        solid_color(0.92, 0.18, 0.16)?,
    )?;
    let second = editor.add_body_shape_with_fill(
        editor.body_text()?.encode_utf16().count(),
        "Originally Front",
        OVERLAP_POSITION,
        OVERLAP_SIZE,
        Preset::Ellipse,
        solid_color(0.12, 0.38, 0.92)?,
    )?;
    editor.move_body_drawable(
        DrawableObjectId::from_object_id(first.drawable_object_id)?,
        DrawableLayerMove::ToFront,
    )?;
    editor.save(output)?;
    println!(
        "moved Pages drawable {} ahead of {}",
        first.drawable_object_id, second.drawable_object_id
    );
    Ok(())
}

fn solid_color(red: f32, green: f32, blue: f32) -> litchi_iwa::Result<ShapeFill> {
    Ok(ShapeFill::Solid(RgbaColor::new(
        red,
        green,
        blue,
        1.0,
        RgbColorSpace::Srgb,
    )?))
}
