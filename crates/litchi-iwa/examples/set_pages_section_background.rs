//! Set a native solid-color Pages section background.

use std::env;

use litchi_iwa::pages::{PagesEditor, PagesSectionBackground};
use litchi_iwa_common::color::{RgbColorSpace, Rgba};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: set_pages_section_background <input.pages> <output.pages> <section-id> \
         <red> <green> <blue> <alpha> <srgb|p3>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let section_id = arguments
        .next()
        .ok_or("missing section ID")?
        .parse::<u64>()?;
    let mut component = |name| -> Result<f32, Box<dyn std::error::Error>> {
        Ok(arguments.next().ok_or(name)?.parse::<f32>()?)
    };
    let red = component("missing red component")?;
    let green = component("missing green component")?;
    let blue = component("missing blue component")?;
    let alpha = component("missing alpha component")?;
    let color_space = match arguments.next().as_deref() {
        Some("srgb") => RgbColorSpace::Srgb,
        Some("p3") => RgbColorSpace::DisplayP3,
        _ => return Err("color space must be srgb or p3".into()),
    };

    let mut editor = PagesEditor::open(input)?;
    editor.set_section_background(
        section_id,
        PagesSectionBackground::Solid(Rgba::new(red, green, blue, alpha, color_space)?),
    )?;
    editor.save(output)?;
    Ok(())
}
