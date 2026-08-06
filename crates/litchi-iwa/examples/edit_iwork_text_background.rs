//! Set one uniform typed solid background in an iWork text storage.

use std::env;

use litchi_iwa::shapes::{RgbColorSpace, RgbaColor};
use litchi_iwa::text::{IWorkTextEditor, Background};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_text_background <input> <output> <storage-id> \
         <none|srgb|display-p3> [red green blue alpha]",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let background = match arguments.next().as_deref() {
        Some("none") => Background::None,
        Some(color_space @ ("srgb" | "display-p3")) => {
            let red = arguments.next().ok_or("missing red component")?.parse()?;
            let green = arguments.next().ok_or("missing green component")?.parse()?;
            let blue = arguments.next().ok_or("missing blue component")?.parse()?;
            let alpha = arguments.next().ok_or("missing alpha component")?.parse()?;
            let color_space = match color_space {
                "srgb" => RgbColorSpace::Srgb,
                "display-p3" => RgbColorSpace::DisplayP3,
                _ => unreachable!("guarded color space"),
            };
            Background::Color(RgbaColor::new(red, green, blue, alpha, color_space)?)
        },
        Some(_) => return Err("background must be none, srgb, or display-p3".into()),
        None => return Err("missing background".into()),
    };
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_background(storage_id, background)?;
    editor.save(output)?;
    Ok(())
}
