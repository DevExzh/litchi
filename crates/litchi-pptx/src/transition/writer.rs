//! `PresentationML` transition writer.

use std::fmt::Write as _;

use crate::{Error, Result};

use super::model::{Axis, Corner, InOut, Kind, Origin, Ripple, Shape, Side, Speed, Transition};
use super::reader::P14;

const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Serializes a transition into a newly allocated XML fragment.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write(value: &Transition) -> Result<String> {
    let mut xml = String::new();
    write_to(value, &mut xml)?;
    Ok(xml)
}

/// Appends a transition XML fragment to an existing text buffer.
///
/// This is the preferred package-writer seam because it avoids an
/// intermediate allocation and copy.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write_to(value: &Transition, xml: &mut String) -> Result<()> {
    validate_raw(value)?;

    if matches!(value.kind, Kind::Ripple(_)) {
        write_alternate_start(xml);
        xml.push_str("<mc:Choice Requires=\"p14\">");
        write_transition(value, xml, value.duration, Effect::Value)?;
        xml.push_str("</mc:Choice><mc:Fallback>");
        write_transition(value, xml, None, Effect::FadeFallback)?;
        xml.push_str("</mc:Fallback></mc:AlternateContent>");
    } else if value.duration.is_some() {
        write_alternate_start(xml);
        xml.push_str("<mc:Choice Requires=\"p14\">");
        write_transition(value, xml, value.duration, Effect::Value)?;
        xml.push_str("</mc:Choice><mc:Fallback>");
        write_transition(value, xml, None, Effect::Value)?;
        xml.push_str("</mc:Fallback></mc:AlternateContent>");
    } else {
        write_transition(value, xml, None, Effect::Value)?;
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum Effect {
    Value,
    FadeFallback,
}

fn write_alternate_start(xml: &mut String) {
    xml.push_str("<mc:AlternateContent xmlns:mc=\"");
    xml.push_str(MCE);
    xml.push_str("\" xmlns:p14=\"");
    xml.push_str(P14);
    xml.push_str("\">");
}

fn write_transition(
    value: &Transition,
    xml: &mut String,
    duration: Option<super::Ms>,
    effect: Effect,
) -> Result<()> {
    xml.push_str("<p:transition spd=\"");
    xml.push_str(speed(value.speed));
    xml.push('"');
    if let Some(duration) = duration {
        write!(xml, " p14:dur=\"{}\"", duration.get()).map_err(|_err| Error::Write)?;
    }
    if !value.click {
        xml.push_str(" advClick=\"0\"");
    }
    if let Some(after) = value.after {
        write!(xml, " advTm=\"{}\"", after.get()).map_err(|_err| Error::Write)?;
    }
    xml.push('>');

    for raw in value.before() {
        write_raw(raw, xml);
    }
    match effect {
        Effect::Value => write_effect(value, xml)?,
        Effect::FadeFallback => xml.push_str("<p:fade/>"),
    }
    for raw in value.after_effect() {
        write_raw(raw, xml);
    }
    xml.push_str("</p:transition>");
    Ok(())
}

fn write_effect(value: &Transition, xml: &mut String) -> Result<()> {
    if let Some(raw) = value.effect_xml() {
        write_raw(raw, xml);
        return Ok(());
    }

    match &value.kind {
        Kind::None => {},
        Kind::Cut { black } => write_black(xml, "cut", *black),
        Kind::Fade { black } => write_black(xml, "fade", *black),
        Kind::Push(side) => write_direction(xml, "push", "dir", side_value(*side)),
        Kind::Wipe(side) => write_direction(xml, "wipe", "dir", side_value(*side)),
        Kind::Split { axis, toward } => {
            xml.push_str("<p:split orient=\"");
            xml.push_str(axis_value(*axis));
            xml.push('"');
            if let Some(toward) = toward {
                xml.push_str(" dir=\"");
                xml.push_str(in_out_value(*toward));
                xml.push('"');
            }
            xml.push_str("/>");
        },
        Kind::Uncover(origin) => {
            write_direction(xml, "pull", "dir", origin_value(*origin));
        },
        Kind::Cover(origin) => {
            write_direction(xml, "cover", "dir", origin_value(*origin));
        },
        Kind::Dissolve => xml.push_str("<p:dissolve/>"),
        Kind::Blinds(axis) => write_direction(xml, "blinds", "dir", axis_value(*axis)),
        Kind::Checker(axis) => write_direction(xml, "checker", "dir", axis_value(*axis)),
        Kind::RandomBars(axis) => write_direction(xml, "randomBar", "dir", axis_value(*axis)),
        Kind::Shape(shape) => match shape {
            Shape::Circle => xml.push_str("<p:circle/>"),
            Shape::Diamond => xml.push_str("<p:diamond/>"),
            Shape::Plus => xml.push_str("<p:plus/>"),
        },
        Kind::Wedge => xml.push_str("<p:wedge/>"),
        Kind::Zoom(direction) => {
            write_direction(xml, "zoom", "dir", in_out_value(*direction));
        },
        Kind::Random => xml.push_str("<p:random/>"),
        Kind::Wheel(spokes) => {
            write!(xml, "<p:wheel spokes=\"{}\"/>", spokes.get()).map_err(|_err| Error::Write)?;
        },
        Kind::Newsflash => xml.push_str("<p:newsflash/>"),
        Kind::Ripple(direction) => {
            xml.push_str("<p14:ripple dir=\"");
            xml.push_str(ripple_value(*direction));
            xml.push_str("\"/>");
        },
        Kind::Strips(corner) => {
            write_direction(xml, "strips", "dir", corner_value(*corner));
        },
        Kind::Comb(axis) => write_direction(xml, "comb", "dir", axis_value(*axis)),
        Kind::Raw(raw) => write_raw(raw, xml),
    }
    Ok(())
}

fn validate_raw(value: &Transition) -> Result<()> {
    let effect = value.effect_xml().or_else(|| match value.kind() {
        Kind::Raw(raw) => Some(raw),
        _ => None,
    });
    let nonportable = effect
        .into_iter()
        .chain(value.before())
        .chain(value.after_effect())
        .any(|raw| !raw.is_portable());
    if nonportable {
        Err(Error::Invalid(
            "retained transition XML depends on a namespace prefix declared outside its subtree"
                .into(),
        ))
    } else {
        Ok(())
    }
}

fn write_raw(raw: &super::Raw, xml: &mut String) {
    xml.push_str(raw.xml());
}

fn write_black(xml: &mut String, tag: &str, black: Option<bool>) {
    xml.push_str("<p:");
    xml.push_str(tag);
    if let Some(black) = black {
        xml.push_str(" thruBlk=\"");
        xml.push_str(if black { "1" } else { "0" });
        xml.push('"');
    }
    xml.push_str("/>");
}

fn write_direction(xml: &mut String, tag: &str, attribute: &str, value: &str) {
    xml.push_str("<p:");
    xml.push_str(tag);
    xml.push(' ');
    xml.push_str(attribute);
    xml.push_str("=\"");
    xml.push_str(value);
    xml.push_str("\"/>");
}

fn speed(value: Speed) -> &'static str {
    match value {
        Speed::Slow => "slow",
        Speed::Medium => "med",
        Speed::Fast => "fast",
    }
}

fn side_value(value: Side) -> &'static str {
    match value {
        Side::Left => "l",
        Side::Right => "r",
        Side::Up => "u",
        Side::Down => "d",
    }
}

fn axis_value(value: Axis) -> &'static str {
    match value {
        Axis::Horizontal => "horz",
        Axis::Vertical => "vert",
    }
}

fn corner_value(value: Corner) -> &'static str {
    match value {
        Corner::LeftUp => "lu",
        Corner::RightUp => "ru",
        Corner::LeftDown => "ld",
        Corner::RightDown => "rd",
    }
}

fn origin_value(value: Origin) -> &'static str {
    match value {
        Origin::Left => "l",
        Origin::Right => "r",
        Origin::Up => "u",
        Origin::Down => "d",
        Origin::LeftUp => "lu",
        Origin::RightUp => "ru",
        Origin::LeftDown => "ld",
        Origin::RightDown => "rd",
    }
}

fn in_out_value(value: InOut) -> &'static str {
    match value {
        InOut::In => "in",
        InOut::Out => "out",
    }
}

fn ripple_value(value: Ripple) -> &'static str {
    match value {
        Ripple::Center => "center",
        Ripple::LeftUp => "lu",
        Ripple::RightUp => "ru",
        Ripple::LeftDown => "ld",
        Ripple::RightDown => "rd",
    }
}
