//! Set one uniform typed RGB color in an iWork text storage.

use std::env;

use litchi_iwa::shapes::{RgbColorSpace, RgbaColor};
use litchi_iwa::text::IWorkTextEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_text_color <input> <output> <storage-id> <red> <green> <blue> <alpha> <srgb|p3>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let red = arguments.next().ok_or("missing red channel")?.parse()?;
    let green = arguments.next().ok_or("missing green channel")?.parse()?;
    let blue = arguments.next().ok_or("missing blue channel")?.parse()?;
    let alpha = arguments.next().ok_or("missing alpha channel")?.parse()?;
    let color_space = match arguments.next().as_deref() {
        Some("srgb") => RgbColorSpace::Srgb,
        Some("p3") => RgbColorSpace::DisplayP3,
        Some(_) => return Err("color space must be srgb or p3".into()),
        None => return Err("missing color space".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let color = RgbaColor::new(red, green, blue, alpha, color_space)?;
    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_color(storage_id, color)?;
    editor.save(output)?;
    Ok(())
}
