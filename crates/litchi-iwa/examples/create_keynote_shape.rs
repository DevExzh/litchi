//! Create a Keynote presentation and editable preset shape without an input package.

use std::env;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeFill, ShapeGradient,
    ShapeGradientAngle, ShapeGradientKind, ShapeGradientOpacity, ShapeGradientStop,
    ShapeGradientStopMidpoint, ShapeGradientStopPosition, ShapePreset,
};

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
    let created = editor.add_slide_shape_with_fill(
        0,
        &text,
        DrawablePoint { x: 720.0, y: 660.0 },
        DrawableSize {
            width: 480.0,
            height: 240.0,
        },
        ShapePreset::RightArrow,
        ShapeFill::Gradient(ShapeGradient::advanced(
            ShapeGradientKind::Radial,
            vec![
                ShapeGradientStop::new(
                    RgbaColor::new(0.98, 0.62, 0.08, 1.0, RgbColorSpace::DisplayP3)?,
                    ShapeGradientStopPosition::START,
                    ShapeGradientStopMidpoint::new(0.4)?,
                ),
                ShapeGradientStop::new(
                    RgbaColor::new(0.72, 0.08, 0.38, 1.0, RgbColorSpace::DisplayP3)?,
                    ShapeGradientStopPosition::END,
                    ShapeGradientStopMidpoint::CENTER,
                ),
            ],
            ShapeGradientOpacity::OPAQUE,
            ShapeGradientAngle::from_degrees(315.0)?,
        )?),
    )?;
    editor.save(output)?;
    println!(
        "created Keynote {:?} {:?} {} with storage {}",
        created.kind, created.preset, created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}
