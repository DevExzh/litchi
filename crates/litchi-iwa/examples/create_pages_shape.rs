//! Create a Pages document and editable preset shape without an input package.

use std::{env, fs, path::Path};

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{
    Appearance, BlurRadius, DrawablePoint, DrawableSize, Drop, Offset, RgbColorSpace, RgbaColor,
    Shadow, ShapeFill, ShapeImageFillTechnique,
};
use litchi_iwa::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
use litchi_iwa_common::shape::effects::{
    Effects, Opacity as EffectsOpacity, Reflection, ReflectionOpacity,
};
use litchi_iwa_common::shape::fill::{Angle, Gradient};
use litchi_iwa_common::shape::path::Preset;
use litchi_iwa_common::shape::shadow::{Angle as ShadowAngle, Opacity as ShadowOpacity};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_shape <output.pages> [text] [fill-image]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Built from typed IWA objects".to_owned());
    let fill_image = arguments.next();
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let body = "Pages shape created entirely by litchi-iwa";
    let mut editor = PagesEditor::create_with_text(body)?;
    let created = editor.add_body_shape_with_fill(
        body.encode_utf16().count(),
        &text,
        DrawablePoint { x: 180.0, y: 240.0 },
        DrawableSize {
            width: 300.0,
            height: 150.0,
        },
        Preset::RightArrow,
        ShapeFill::Gradient(Gradient::linear(
            RgbaColor::new(0.88, 0.18, 0.12, 1.0, RgbColorSpace::DisplayP3)?,
            RgbaColor::new(0.98, 0.65, 0.08, 1.0, RgbColorSpace::DisplayP3)?,
            Angle::from_degrees(45.0)?,
        )),
    )?;
    if let Some(path) = fill_image {
        let filename = preferred_filename(&path)?;
        editor.set_body_shape_image_fill(
            created.drawable_object_id,
            filename,
            &fs::read(&path)?,
            ShapeImageFillTechnique::ScaleToFill,
            None,
        )?;
    }
    editor.set_body_shape_effects(
        created.drawable_object_id,
        Effects::new(
            EffectsOpacity::new(0.72)?,
            Reflection::Enabled(ReflectionOpacity::new(0.35)?),
        ),
    )?;
    editor.set_body_shape_shadow(
        created.drawable_object_id,
        Shadow::Drop(Drop::new(
            Appearance::new(
                RgbaColor::black(),
                BlurRadius::from_points(7)?,
                Offset::from_points(11.0)?,
                ShadowOpacity::new(0.42)?,
            ),
            ShadowAngle::from_degrees(135.0)?,
        )),
    )?;
    editor.set_body_shape_text_layout(
        created.drawable_object_id,
        Layout::new(
            VerticalAlignment::Middle,
            Insets::uniform(Inset::from_points(12.0)?),
            AutoSize::Fixed,
        ),
    )?;
    editor.set_body_shape_title(created.drawable_object_id, "Typed Pages shape")?;
    editor.set_body_shape_caption(
        created.drawable_object_id,
        "Native title and caption created by litchi-iwa",
    )?;
    editor.save(output)?;
    println!(
        "created Pages {:?} {:?} {} with storage {} at UTF-16 anchor {}",
        created.kind,
        created.preset,
        created.drawable_object_id,
        created.storage.id,
        created.anchor_character_index
    );
    Ok(())
}

fn preferred_filename(path: &str) -> Result<&str, Box<dyn std::error::Error>> {
    Path::new(path)
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| "fill image path has no UTF-8 filename".into())
}
