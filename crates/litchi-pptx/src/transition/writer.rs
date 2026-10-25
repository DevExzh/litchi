//! `PresentationML` transition writer.

use crate::time::Offset;
use crate::{Error, Result};

use super::model::{
    Axis, Corner, InOut, Kind, LeftRight, Origin, Ripple, Shape, Side, Speed, Transition,
};
use super::reader::{Limits, P14, P15, P159};

const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Serializes a transition into a newly allocated XML fragment.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write(value: &Transition, limits: Limits) -> Result<String> {
    let mut xml = String::new();
    write_to(value, &mut xml, limits)?;
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
pub fn write_to(value: &Transition, xml: &mut String, limits: Limits) -> Result<()> {
    validate_raw(value)?;

    // Count the exact fragment before touching the destination.  In
    // particular, retained XML and escaped preset names are data-dependent;
    // a point-count-style estimate cannot be a safe reservation for them.
    let mut count = CountingXml::default();
    emit(value, &mut count)?;
    let output_len = count.finish()?;
    if output_len > limits.output_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: output_len,
            maximum: limits.output_bytes(),
        });
    }

    let start = xml.len();
    xml.try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "transition XML",
            source,
        })?;
    let result = {
        let mut output = LimitedXml {
            xml,
            start,
            maximum: output_len,
            exceeded: false,
        };
        let result = emit(value, &mut output);
        result.and_then(|()| {
            if output.exceeded {
                Err(Error::Limit {
                    kind: crate::opened::LimitKind::Codec(
                        crate::opened::CodecLimitKind::OutputBytes,
                    ),
                    observed: output_len,
                    maximum: limits.output_bytes(),
                })
            } else if output.xml.len().saturating_sub(start) != output_len {
                Err(Error::Invalid(
                    "transition output length changed during emission".into(),
                ))
            } else {
                Ok(())
            }
        })
    };
    if result.is_err() {
        xml.truncate(start);
    }
    result
}

struct LimitedXml<'a> {
    xml: &'a mut String,
    start: usize,
    maximum: usize,
    exceeded: bool,
}

trait XmlSink: std::fmt::Write {
    fn push_str(&mut self, value: &str);
    fn push(&mut self, value: char);
}

#[derive(Default)]
struct CountingXml {
    length: usize,
    overflowed: bool,
}

impl CountingXml {
    fn finish(self) -> Result<usize> {
        if self.overflowed {
            Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        } else {
            Ok(self.length)
        }
    }
}

impl XmlSink for CountingXml {
    fn push_str(&mut self, value: &str) {
        if let Some(length) = self.length.checked_add(value.len()) {
            self.length = length;
        } else {
            self.overflowed = true;
        }
    }

    fn push(&mut self, value: char) {
        let mut encoded = [0u8; 4];
        self.push_str(value.encode_utf8(&mut encoded));
    }
}

impl std::fmt::Write for CountingXml {
    fn write_str(&mut self, value: &str) -> std::fmt::Result {
        self.push_str(value);
        Ok(())
    }
}

impl LimitedXml<'_> {
    fn push_str(&mut self, value: &str) {
        if self.exceeded {
            return;
        }
        let Some(length) = self
            .xml
            .len()
            .checked_sub(self.start)
            .and_then(|length| length.checked_add(value.len()))
        else {
            self.exceeded = true;
            return;
        };
        let Some(end) = self.xml.len().checked_add(value.len()) else {
            self.exceeded = true;
            return;
        };
        if length > self.maximum || end > self.xml.capacity() {
            self.exceeded = true;
            return;
        }
        self.xml.push_str(value);
    }

    fn push(&mut self, value: char) {
        let mut encoded = [0u8; 4];
        self.push_str(value.encode_utf8(&mut encoded));
    }
}

impl XmlSink for LimitedXml<'_> {
    fn push_str(&mut self, value: &str) {
        Self::push_str(self, value);
    }

    fn push(&mut self, value: char) {
        Self::push(self, value);
    }
}

impl std::fmt::Write for LimitedXml<'_> {
    fn write_str(&mut self, value: &str) -> std::fmt::Result {
        let was_exceeded = self.exceeded;
        self.push_str(value);
        if was_exceeded || self.exceeded {
            Err(std::fmt::Error)
        } else {
            Ok(())
        }
    }
}

fn extension_requirement(value: &Transition) -> Option<&'static str> {
    match (&value.kind, value.duration_offset().is_some()) {
        (kind, _) if is_p14_effect(kind) => Some("p14"),
        (Kind::Morph(_), true) => Some("p14 p159"),
        (Kind::Morph(_), false) => Some("p159"),
        (Kind::Preset(_), true) => Some("p14 p15"),
        (Kind::Preset(_), false) => Some("p15"),
        (_, true) => Some("p14"),
        (_, false) => None,
    }
}

fn is_p14_effect(kind: &Kind) -> bool {
    matches!(
        kind,
        Kind::Ripple(_)
            | Kind::Conveyor(_)
            | Kind::Doors(_)
            | Kind::Ferris(_)
            | Kind::Flash
            | Kind::Flip(_)
            | Kind::FlyThrough(_)
            | Kind::Gallery(_)
            | Kind::Glitter(_)
            | Kind::Honeycomb
            | Kind::Pan(_)
            | Kind::Prism(_)
            | Kind::Reveal(_)
            | Kind::Shred(_)
            | Kind::Switch(_)
            | Kind::Vortex(_)
            | Kind::Warp(_)
            | Kind::WheelReverse(_)
            | Kind::Window(_)
    )
}

fn uses_fade_fallback(kind: &Kind) -> bool {
    is_p14_effect(kind) || matches!(kind, Kind::Morph(_) | Kind::Preset(_))
}

#[derive(Debug, Clone, Copy)]
enum Effect {
    Value,
    FadeFallback,
}

fn emit<S: XmlSink>(value: &Transition, output: &mut S) -> Result<()> {
    if let Some(requires) = extension_requirement(value) {
        write_alternate_start(output, requires);
        output.push_str("<mc:Choice Requires=\"");
        output.push_str(requires);
        output.push_str("\">");
        write_transition(value, output, value.duration_offset(), Effect::Value)?;
        output.push_str("</mc:Choice><mc:Fallback>");
        let fallback = if uses_fade_fallback(&value.kind) {
            Effect::FadeFallback
        } else {
            Effect::Value
        };
        write_transition(value, output, None, fallback)?;
        output.push_str("</mc:Fallback></mc:AlternateContent>");
    } else {
        write_transition(value, output, None, Effect::Value)?;
    }
    Ok(())
}

fn write_alternate_start<S: XmlSink>(xml: &mut S, requires: &str) {
    xml.push_str("<mc:AlternateContent xmlns:mc=\"");
    xml.push_str(MCE);
    if requires
        .split_ascii_whitespace()
        .any(|value| value == "p14")
    {
        xml.push_str("\" xmlns:p14=\"");
        xml.push_str(P14);
    }
    if requires
        .split_ascii_whitespace()
        .any(|value| value == "p15")
    {
        xml.push_str("\" xmlns:p15=\"");
        xml.push_str(P15);
    }
    if requires
        .split_ascii_whitespace()
        .any(|value| value == "p159")
    {
        xml.push_str("\" xmlns:p159=\"");
        xml.push_str(P159);
    }
    xml.push_str("\">");
}

fn write_transition<S: XmlSink>(
    value: &Transition,
    xml: &mut S,
    duration: Option<&Offset>,
    effect: Effect,
) -> Result<()> {
    xml.push_str("<p:transition spd=\"");
    xml.push_str(speed(value.speed));
    xml.push('"');
    if let Some(duration) = duration {
        write!(xml, " p14:dur=\"{}\"", duration.as_str()).map_err(|_err| Error::Write)?;
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

fn write_effect<S: XmlSink>(value: &Transition, xml: &mut S) -> Result<()> {
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
        Kind::Conveyor(direction) => {
            write_p14_left_right(xml, "conveyor", left_right_value(*direction));
        },
        Kind::Doors(axis) => write_p14_direction(xml, "doors", axis_value(*axis)),
        Kind::Ferris(direction) => {
            write_p14_left_right(xml, "ferris", left_right_value(*direction));
        },
        Kind::Flash => xml.push_str("<p14:flash/>"),
        Kind::Flip(direction) => {
            write_p14_left_right(xml, "flip", left_right_value(*direction));
        },
        Kind::FlyThrough(value) => {
            xml.push_str("<p14:flythrough dir=\"");
            xml.push_str(in_out_value(value.direction()));
            xml.push('"');
            if value.bounce() {
                xml.push_str(" hasBounce=\"1\"");
            }
            xml.push_str("/>");
        },
        Kind::Gallery(direction) => {
            write_p14_left_right(xml, "gallery", left_right_value(*direction));
        },
        Kind::Glitter(value) => {
            xml.push_str("<p14:glitter dir=\"");
            xml.push_str(side_value(value.direction()));
            xml.push_str("\" pattern=\"");
            xml.push_str(value.pattern().wire());
            xml.push_str("\"/>");
        },
        Kind::Honeycomb => xml.push_str("<p14:honeycomb/>"),
        Kind::Pan(direction) => {
            write_p14_direction(xml, "pan", side_value(*direction));
        },
        Kind::Prism(value) => {
            xml.push_str("<p14:prism dir=\"");
            xml.push_str(side_value(value.direction()));
            xml.push('"');
            if value.content() {
                xml.push_str(" isContent=\"1\"");
            }
            if value.inverted() {
                xml.push_str(" isInverted=\"1\"");
            }
            xml.push_str("/>");
        },
        Kind::Reveal(value) => {
            xml.push_str("<p14:reveal");
            if let Some(direction) = left_right_value(value.direction()) {
                xml.push_str(" dir=\"");
                xml.push_str(direction);
                xml.push('"');
            }
            if value.through_black() {
                xml.push_str(" thruBlk=\"1\"");
            }
            xml.push_str("/>");
        },
        Kind::Shred(value) => {
            xml.push_str("<p14:shred pattern=\"");
            xml.push_str(value.pattern().wire());
            xml.push_str("\" dir=\"");
            xml.push_str(in_out_value(value.direction()));
            xml.push_str("\"/>");
        },
        Kind::Switch(direction) => {
            write_p14_left_right(xml, "switch", left_right_value(*direction));
        },
        Kind::Vortex(direction) => {
            write_p14_direction(xml, "vortex", side_value(*direction));
        },
        Kind::Warp(direction) => {
            write_p14_direction(xml, "warp", in_out_value(*direction));
        },
        Kind::WheelReverse(spokes) => {
            write!(xml, "<p14:wheelReverse spokes=\"{}\"/>", spokes.get())
                .map_err(|_err| Error::Write)?;
        },
        Kind::Window(axis) => {
            write_p14_direction(xml, "window", axis_value(*axis));
        },
        Kind::Morph(option) => {
            xml.push_str("<p159:morph option=\"");
            xml.push_str(option.wire());
            xml.push_str("\"/>");
        },
        Kind::Preset(preset) => write_preset(preset, xml)?,
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

fn write_raw<S: XmlSink>(raw: &super::Raw, xml: &mut S) {
    xml.push_str(raw.xml());
}

fn write_preset<S: XmlSink>(value: &super::Preset, xml: &mut S) -> Result<()> {
    xml.push_str("<p15:prstTrans");
    if let Some(name) = value.name() {
        xml.push_str(" prst=\"");
        escape_attribute(name, xml)?;
        xml.push('"');
    }
    if value.invert_x() {
        xml.push_str(" invX=\"1\"");
    }
    if value.invert_y() {
        xml.push_str(" invY=\"1\"");
    }
    xml.push_str("/>");
    Ok(())
}

fn escape_attribute<S: XmlSink>(value: &str, output: &mut S) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
            character
                if {
                    let value = character as u32;
                    !matches!(
                        value,
                        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
                    )
                } =>
            {
                return Err(Error::Invalid(
                    "preset transition name contains an invalid XML character".into(),
                ));
            },
            character => output.push(character),
        }
    }
    Ok(())
}

fn write_black<S: XmlSink>(xml: &mut S, tag: &str, black: Option<bool>) {
    xml.push_str("<p:");
    xml.push_str(tag);
    if let Some(black) = black {
        xml.push_str(" thruBlk=\"");
        xml.push_str(if black { "1" } else { "0" });
        xml.push('"');
    }
    xml.push_str("/>");
}

fn write_direction<S: XmlSink>(xml: &mut S, tag: &str, attribute: &str, value: &str) {
    xml.push_str("<p:");
    xml.push_str(tag);
    xml.push(' ');
    xml.push_str(attribute);
    xml.push_str("=\"");
    xml.push_str(value);
    xml.push_str("\"/>");
}

fn write_p14_direction<S: XmlSink>(xml: &mut S, tag: &str, value: &str) {
    xml.push_str("<p14:");
    xml.push_str(tag);
    xml.push_str(" dir=\"");
    xml.push_str(value);
    xml.push_str("\"/>");
}

fn write_p14_left_right<S: XmlSink>(xml: &mut S, tag: &str, value: Option<&str>) {
    xml.push_str("<p14:");
    xml.push_str(tag);
    if let Some(value) = value {
        xml.push_str(" dir=\"");
        xml.push_str(value);
        xml.push('"');
    }
    xml.push_str("/>");
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

fn left_right_value(value: LeftRight) -> Option<&'static str> {
    value.wire()
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
