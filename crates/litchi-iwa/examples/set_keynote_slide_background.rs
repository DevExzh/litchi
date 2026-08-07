//! Set one Keynote slide background without automating Keynote.

use std::env;

use litchi_iwa::keynote::{
    Angle, Background, Gradient, KeynoteEditor, Kind, RgbColorSpace, Rgba, Stop,
};
use litchi_iwa_common::shape::fill::{Opacity, StopMidpoint, StopPosition};

enum Operation {
    Set(Background),
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
        ("none", []) => Operation::Set(Background::None),
        ("solid", [red, green, blue, alpha, color_space]) => {
            Operation::Set(Background::Solid(Rgba::new(
                red.parse()?,
                green.parse()?,
                blue.parse()?,
                alpha.parse()?,
                parse_color_space(color_space)?,
            )?))
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
        ) => Operation::Set(Background::Gradient(Gradient::linear(
            parse_color(start_red, start_green, start_blue, start_alpha, start_space)?,
            parse_color(end_red, end_green, end_blue, end_alpha, end_space)?,
            Angle::from_degrees(angle.parse()?)?,
        ))),
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
            Operation::Set(Background::Gradient(Gradient::advanced(
                Kind::Radial,
                vec![
                    Stop::new(start, StopPosition::START, StopMidpoint::CENTER),
                    Stop::new(end, StopPosition::END, StopMidpoint::CENTER),
                ],
                Opacity::OPAQUE,
                Angle::from_degrees(angle.parse()?)?,
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
) -> Result<Rgba, Box<dyn std::error::Error>> {
    Ok(Rgba::new(
        red.parse()?,
        green.parse()?,
        blue.parse()?,
        alpha.parse()?,
        parse_color_space(color_space)?,
    )?)
}

fn parse_color_space(value: &str) -> Result<RgbColorSpace, Box<dyn std::error::Error>> {
    match value {
        "srgb" => Ok(RgbColorSpace::Srgb),
        "display-p3" => Ok(RgbColorSpace::DisplayP3),
        _ => Err(usage().into()),
    }
}
