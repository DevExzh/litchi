//! Create a Keynote presentation and editable preset shape without an input package.

use std::{env, fs, path::Path};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeCurvedShadow, ShapeFill,
    ShapeImageFillTechnique, ShapeShadow, ShapeShadowAngle, ShapeShadowAppearance,
    ShapeShadowBlurRadius, ShapeShadowCurve, ShapeShadowOffset, ShapeShadowOpacity,
};
use litchi_iwa::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
use litchi_iwa_common::shape::effects::{Effects, Opacity, Reflection, ReflectionOpacity};
use litchi_iwa_common::shape::fill::{
    Angle, Gradient, Kind, Opacity as GradientOpacity, Stop, StopMidpoint, StopPosition,
};
use litchi_iwa_common::shape::path::Preset;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_shape <output.key> [text] [fill-image]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Built from typed IWA objects".to_owned());
    let fill_image = arguments.next();
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
        Preset::RightArrow,
        ShapeFill::Gradient(Gradient::advanced(
            Kind::Radial,
            vec![
                Stop::new(
                    RgbaColor::new(0.98, 0.62, 0.08, 1.0, RgbColorSpace::DisplayP3)?,
                    StopPosition::START,
                    StopMidpoint::new(0.4)?,
                ),
                Stop::new(
                    RgbaColor::new(0.72, 0.08, 0.38, 1.0, RgbColorSpace::DisplayP3)?,
                    StopPosition::END,
                    StopMidpoint::CENTER,
                ),
            ],
            GradientOpacity::OPAQUE,
            Angle::from_degrees(315.0)?,
        )?),
    )?;
    if let Some(path) = fill_image {
        let filename = preferred_filename(&path)?;
        editor.set_slide_shape_image_fill(
            0,
            created.drawable_object_id,
            filename,
            &fs::read(&path)?,
            ShapeImageFillTechnique::Tile,
            Some(RgbaColor::new(0.0, 0.0, 0.0, 0.5, RgbColorSpace::Srgb)?),
        )?;
    }
    editor.set_slide_shape_effects(
        0,
        created.drawable_object_id,
        Effects::new(
            Opacity::new(0.61)?,
            Reflection::Enabled(ReflectionOpacity::new(0.2)?),
        ),
    )?;
    editor.set_slide_shape_shadow(
        0,
        created.drawable_object_id,
        ShapeShadow::Curved(ShapeCurvedShadow::new(
            ShapeShadowAppearance::new(
                RgbaColor::black(),
                ShapeShadowBlurRadius::from_points(15)?,
                ShapeShadowOffset::from_points(4.0)?,
                ShapeShadowOpacity::new(0.73)?,
            ),
            ShapeShadowAngle::from_degrees(310.0)?,
            ShapeShadowCurve::new(0.2)?,
        )),
    )?;
    editor.set_slide_shape_text_layout(
        0,
        created.drawable_object_id,
        Layout::new(
            VerticalAlignment::Middle,
            Insets::uniform(Inset::from_points(14.0)?),
            AutoSize::ShrinkToFit,
        ),
    )?;
    editor.set_slide_shape_title(0, created.drawable_object_id, "Typed Keynote shape")?;
    editor.set_slide_shape_caption(
        0,
        created.drawable_object_id,
        "Native title and caption created by litchi-iwa",
    )?;
    editor.save(output)?;
    println!(
        "created Keynote {:?} {:?} {} with storage {}",
        created.kind, created.preset, created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}

fn preferred_filename(path: &str) -> Result<&str, Box<dyn std::error::Error>> {
    Path::new(path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| "fill image path has no UTF-8 filename".into())
}
