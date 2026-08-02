//! Set one Keynote slide background without automating Keynote.

use std::env;

use litchi_iwa::keynote::{
    KeynoteEditor, KeynoteGradient, KeynoteGradientAngle, KeynoteGradientKind, KeynoteGradientStop,
    KeynoteRgbColorSpace, KeynoteRgbaColor, KeynoteSlideBackground,
};

enum Operation {
    Set(KeynoteSlideBackground),
    Reset,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments = env::args().skip(1).collect::<Vec<_>>();
    let [input, output, slide_index, mode, rest @ ..] = arguments.as_slice() else {
        return Err(usage().into());
    };
    let slide_index = slide_index.parse::<usize>()?;
    let operation = match (mode.as_str(), rest) {
        ("reset", []) => Operation::Reset,
        ("none", []) => Operation::Set(KeynoteSlideBackground::None),
        ("solid", [red, green, blue, alpha, color_space]) => {
            Operation::Set(KeynoteSlideBackground::Solid(KeynoteRgbaColor {
                red: red.parse()?,
                green: green.parse()?,
                blue: blue.parse()?,
                alpha: alpha.parse()?,
                color_space: parse_color_space(color_space)?,
            }))
        },
        (
            "linear-gradient",
            [
                angle,
                start_red,
                start_green,
                start_blue,
                start_alpha,
                start_space,
                end_red,
                end_green,
                end_blue,
                end_alpha,
                end_space,
            ],
        ) => Operation::Set(KeynoteSlideBackground::Gradient(KeynoteGradient::linear(
            parse_color(start_red, start_green, start_blue, start_alpha, start_space)?,
            parse_color(end_red, end_green, end_blue, end_alpha, end_space)?,
            KeynoteGradientAngle::from_degrees(angle.parse()?)?,
        )?)),
        (
            "radial-gradient",
            [
                angle,
                start_red,
                start_green,
                start_blue,
                start_alpha,
                start_space,
                end_red,
                end_green,
                end_blue,
                end_alpha,
                end_space,
            ],
        ) => {
            let start = parse_color(start_red, start_green, start_blue, start_alpha, start_space)?;
            let end = parse_color(end_red, end_green, end_blue, end_alpha, end_space)?;
            Operation::Set(KeynoteSlideBackground::Gradient(KeynoteGradient::advanced(
                KeynoteGradientKind::Radial,
                vec![
                    KeynoteGradientStop::new(start, 0.0, 0.5)?,
                    KeynoteGradientStop::new(end, 1.0, 0.5)?,
                ],
                1.0,
                KeynoteGradientAngle::from_degrees(angle.parse()?)?,
            )?))
        },
        _ => return Err(usage().into()),
    };

    let mut editor = KeynoteEditor::open(input)?;
    match operation {
        Operation::Set(background) => editor.set_slide_background(slide_index, background)?,
        Operation::Reset => {
            editor.reset_slide_background(slide_index)?;
        },
    }
    editor.save(output)?;
    Ok(())
}

fn usage() -> &'static str {
    "usage: set_keynote_slide_background <input.key> <output.key> <zero-based-slide-index> \
<reset|none|solid <red> <green> <blue> <alpha> <srgb|display-p3>|\
<linear-gradient|radial-gradient> <angle-degrees> <start-red> <start-green> <start-blue> <start-alpha> \
<srgb|display-p3> <end-red> <end-green> <end-blue> <end-alpha> <srgb|display-p3>>"
}

fn parse_color(
    red: &str,
    green: &str,
    blue: &str,
    alpha: &str,
    color_space: &str,
) -> Result<KeynoteRgbaColor, Box<dyn std::error::Error>> {
    Ok(KeynoteRgbaColor {
        red: red.parse()?,
        green: green.parse()?,
        blue: blue.parse()?,
        alpha: alpha.parse()?,
        color_space: parse_color_space(color_space)?,
    })
}

fn parse_color_space(value: &str) -> Result<KeynoteRgbColorSpace, Box<dyn std::error::Error>> {
    match value {
        "srgb" => Ok(KeynoteRgbColorSpace::Srgb),
        "display-p3" => Ok(KeynoteRgbColorSpace::DisplayP3),
        _ => Err(usage().into()),
    }
}
