//! Set one Keynote slide background without automating Keynote.

use std::env;

use litchi_iwa::keynote::{
    KeynoteEditor, KeynoteRgbColorSpace, KeynoteRgbaColor, KeynoteSlideBackground,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    let [input, output, slide_index, mode, rest @ ..] = arguments.as_slice() else {
        return Err(usage().into());
    };
    let slide_index = slide_index.parse::<usize>()?;
    let background = match (mode.as_str(), rest) {
        ("none", []) => KeynoteSlideBackground::None,
        ("solid", [red, green, blue, alpha, color_space]) => {
            let color_space = match color_space.as_str() {
                "srgb" => KeynoteRgbColorSpace::Srgb,
                "display-p3" => KeynoteRgbColorSpace::DisplayP3,
                _ => return Err(usage().into()),
            };
            KeynoteSlideBackground::Solid(KeynoteRgbaColor {
                red: red.parse()?,
                green: green.parse()?,
                blue: blue.parse()?,
                alpha: alpha.parse()?,
                color_space,
            })
        },
        _ => return Err(usage().into()),
    };

    let mut editor = KeynoteEditor::open(input)?;
    editor.set_slide_background(slide_index, background)?;
    editor.save(output)?;
    Ok(())
}

fn usage() -> &'static str {
    "usage: set_keynote_slide_background <input.key> <output.key> <zero-based-slide-index> \
<none|solid <red> <green> <blue> <alpha> <srgb|display-p3>>"
}
