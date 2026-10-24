//! Bounded XML codec for `DrawingML` color fragments.

use std::fmt;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::{Error, Result};

use super::model::{
    Angle, Base, FixedPercentage, Hsl, PositiveAngle, PositiveFixedPercentage, PositivePercentage,
    Preset, Rgb, ScRgb, Scheme, System, Transform, Transformed, Unknown, Value,
};
use super::validation;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

pub use super::validation::{Limits, MAX_DEPTH, MAX_NODES, MAX_TRANSFORMS, MAX_XML_BYTES};

const VAL: &[u8] = b"val";
const LAST_CLR: &[u8] = b"lastClr";
const RED: &[u8] = b"r";
const GREEN: &[u8] = b"g";
const BLUE: &[u8] = b"b";
const HUE: &[u8] = b"hue";
const SATURATION: &[u8] = b"sat";
const LUMINANCE: &[u8] = b"lum";

#[derive(Debug, Clone, Copy)]
enum ChoiceKind {
    Rgb,
    Scheme,
    ScRgb,
    Hsl,
    System,
    Preset,
}

#[derive(Debug, Clone, Copy)]
enum BaseRef<'a> {
    Rgb(Rgb),
    Scheme(Scheme),
    ScRgb(ScRgb),
    Hsl(Hsl),
    System(&'a System),
    Preset(&'a Preset),
}

impl<'a> From<&'a Base> for BaseRef<'a> {
    fn from(value: &'a Base) -> Self {
        match value {
            Base::Rgb(value) => Self::Rgb(*value),
            Base::Scheme(value) => Self::Scheme(*value),
            Base::ScRgb(value) => Self::ScRgb(*value),
            Base::Hsl(value) => Self::Hsl(*value),
            Base::System(value) => Self::System(value),
            Base::Preset(value) => Self::Preset(value),
        }
    }
}

/// Read one `DrawingML` color-choice fragment.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn read(xml: &[u8], limits: &Limits) -> Result<Value> {
    let limits = limits.validate()?;
    let validated = validation::validated_fragment(xml, limits)?;
    let mut reader = Reader::from_reader(validated);

    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Text(text)
                if text
                    .decode()
                    .map_err(|error| xml_error(&error))?
                    .trim()
                    .is_empty() => {},
            Event::Start(element) => {
                return Ok(read_choice(&mut reader, &element, limits.max_transforms)?
                    .unwrap_or_else(|| Value::Unknown(Unknown::from_validated(validated))));
            },
            Event::Empty(element) => {
                return Ok(
                    parse_empty_choice(&mut reader, &element, limits.max_transforms)?
                        .unwrap_or_else(|| Value::Unknown(Unknown::from_validated(validated))),
                );
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::Eof => return Ok(Value::Unknown(Unknown::from_validated(validated))),
        }
    }
}

/// Write one `DrawingML` color-choice fragment using the conventional `a`
/// prefix. The result is a fragment and intentionally has no namespace
/// declaration so the host can retain its own namespace spelling.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write(value: &Value, limits: &Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    if let Value::Unknown(value) = value {
        let validated = validation::validated_fragment(value.as_xml(), limits)?;
        if validated.len() > limits.max_xml_bytes {
            return Err(Error::limit(
                crate::LimitKind::XmlBytes,
                validated.len(),
                limits.max_xml_bytes,
            ));
        }
        return Ok(validated.to_vec());
    }
    validate_typed_limits(value, limits)?;

    let mut length = LengthWriter::default();
    write_value_xml(&mut length, value).map_err(format_error)?;
    if length.overflow {
        return Err(Error::limit(
            crate::LimitKind::XmlBytes,
            usize::MAX,
            limits.max_xml_bytes,
        ));
    }
    if length.len > limits.max_xml_bytes {
        return Err(Error::limit(
            crate::LimitKind::XmlBytes,
            length.len,
            limits.max_xml_bytes,
        ));
    }
    let mut output = String::with_capacity(length.len);
    write_value_xml(&mut output, value).map_err(format_error)?;
    Ok(output.into_bytes())
}

fn read_choice(
    reader: &mut Reader<&[u8]>,
    element: &BytesStart<'_>,
    max_transforms: usize,
) -> Result<Option<Value>> {
    let Some(base) = parse_base(element, reader.decoder())? else {
        return Ok(None);
    };
    let root_name = element.name().as_ref().to_vec();
    let mut transforms = Vec::new();

    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Text(text)
                if text
                    .decode()
                    .map_err(|error| xml_error(&error))?
                    .trim()
                    .is_empty() => {},
            Event::Empty(element) => {
                let Some(transform) = parse_transform(&element, reader.decoder())? else {
                    return Ok(None);
                };
                push_transform(&mut transforms, transform, max_transforms)?;
            },
            Event::Start(element) => {
                let Some(transform) = parse_started_transform(reader, &element)? else {
                    return Ok(None);
                };
                push_transform(&mut transforms, transform, max_transforms)?;
            },
            Event::End(element) if element.name().as_ref() == root_name.as_slice() => break,
            Event::Eof => {
                return Err(Error::Invalid(
                    "DrawingML color choice ended before its root element".into(),
                ));
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => return Ok(None),
        }
    }

    if !tail_is_empty(reader)? {
        return Ok(None);
    }
    if transforms.is_empty() {
        return Ok(Some(Value::from_base(base)));
    }
    Ok(Some(Value::Transformed(Transformed::new(
        base, transforms,
    )?)))
}

fn parse_empty_choice(
    reader: &mut Reader<&[u8]>,
    element: &BytesStart<'_>,
    _max_transforms: usize,
) -> Result<Option<Value>> {
    let decoder = reader.decoder();
    let Some(base) = parse_base(element, decoder)? else {
        return Ok(None);
    };
    if !tail_is_empty(reader)? {
        return Ok(None);
    }
    Ok(Some(Value::from_base(base)))
}

fn parse_base(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<Base>> {
    let kind = match element.local_name().as_ref() {
        b"srgbClr" => ChoiceKind::Rgb,
        b"schemeClr" => ChoiceKind::Scheme,
        b"scrgbClr" => ChoiceKind::ScRgb,
        b"hslClr" => ChoiceKind::Hsl,
        b"sysClr" => ChoiceKind::System,
        b"prstClr" => ChoiceKind::Preset,
        _ => return Ok(None),
    };
    let allowed = match kind {
        ChoiceKind::Rgb | ChoiceKind::Scheme | ChoiceKind::Preset => &[VAL][..],
        ChoiceKind::ScRgb => &[RED, GREEN, BLUE][..],
        ChoiceKind::Hsl => &[HUE, SATURATION, LUMINANCE][..],
        ChoiceKind::System => &[VAL, LAST_CLR][..],
    };
    let Some(attributes) = attributes(element, decoder, allowed)? else {
        return Ok(None);
    };

    match kind {
        ChoiceKind::Rgb => {
            let value = required(&attributes, VAL, "srgbClr")?;
            Ok(Some(Base::Rgb(Rgb::parse(value)?)))
        },
        ChoiceKind::Scheme => {
            Ok(Scheme::from_token(required(&attributes, VAL, "schemeClr")?).map(Base::Scheme))
        },
        ChoiceKind::ScRgb => {
            let red = PositiveFixedPercentage::parse(required(&attributes, RED, "scrgbClr")?)?;
            let green = PositiveFixedPercentage::parse(required(&attributes, GREEN, "scrgbClr")?)?;
            let blue = PositiveFixedPercentage::parse(required(&attributes, BLUE, "scrgbClr")?)?;
            Ok(Some(Base::ScRgb(ScRgb::from_values(red, green, blue))))
        },
        ChoiceKind::Hsl => {
            let hue = PositiveAngle::parse(required(&attributes, HUE, "hslClr")?)?;
            let saturation =
                super::model::Percentage::parse(required(&attributes, SATURATION, "hslClr")?)?;
            let luminance =
                super::model::Percentage::parse(required(&attributes, LUMINANCE, "hslClr")?)?;
            Ok(Some(Base::Hsl(Hsl::from_values(
                hue, saturation, luminance,
            ))))
        },
        ChoiceKind::System => {
            let last = optional(&attributes, LAST_CLR)
                .map(Rgb::parse)
                .transpose()?;
            Ok(System::new(required(&attributes, VAL, "sysClr")?, last)
                .ok()
                .map(Base::System))
        },
        ChoiceKind::Preset => Ok(Preset::new(required(&attributes, VAL, "prstClr")?)
            .ok()
            .map(Base::Preset)),
    }
}

fn parse_started_transform(
    reader: &mut Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<Transform>> {
    let Some(transform) = parse_transform(element, reader.decoder())? else {
        return Ok(None);
    };
    let name = element.name().as_ref().to_vec();
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Text(text)
                if text
                    .decode()
                    .map_err(|error| xml_error(&error))?
                    .trim()
                    .is_empty() => {},
            Event::End(end) if end.name().as_ref() == name.as_slice() => {
                return Ok(Some(transform));
            },
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::Eof => return Ok(None),
        }
    }
}

fn parse_transform(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<Transform>> {
    let local = element.local_name().as_ref().to_vec();
    let (kind, allowed) = match local.as_slice() {
        b"alpha" => (0_u8, &[VAL][..]),
        b"alphaMod" => (1, &[VAL][..]),
        b"alphaOff" => (2, &[VAL][..]),
        b"blue" => (3, &[VAL][..]),
        b"blueMod" => (4, &[VAL][..]),
        b"blueOff" => (5, &[VAL][..]),
        b"comp" => (6, &[][..]),
        b"gamma" => (7, &[][..]),
        b"gray" => (8, &[][..]),
        b"green" => (9, &[VAL][..]),
        b"greenMod" => (10, &[VAL][..]),
        b"greenOff" => (11, &[VAL][..]),
        b"hue" => (12, &[VAL][..]),
        b"hueMod" => (13, &[VAL][..]),
        b"hueOff" => (14, &[VAL][..]),
        b"inv" => (15, &[][..]),
        b"invGamma" => (16, &[][..]),
        b"lum" => (17, &[VAL][..]),
        b"lumMod" => (18, &[VAL][..]),
        b"lumOff" => (19, &[VAL][..]),
        b"red" => (20, &[VAL][..]),
        b"redMod" => (21, &[VAL][..]),
        b"redOff" => (22, &[VAL][..]),
        b"sat" => (23, &[VAL][..]),
        b"satMod" => (24, &[VAL][..]),
        b"satOff" => (25, &[VAL][..]),
        b"shade" => (26, &[VAL][..]),
        b"tint" => (27, &[VAL][..]),
        _ => return Ok(None),
    };
    let Some(attributes) = attributes(element, decoder, allowed)? else {
        return Ok(None);
    };
    let value = || required(&attributes, VAL, &String::from_utf8_lossy(&local));
    Ok(Some(match kind {
        0 => Transform::Alpha(PositiveFixedPercentage::parse(value()?)?),
        1 => Transform::AlphaMod(PositivePercentage::parse(value()?)?),
        2 => Transform::AlphaOff(FixedPercentage::parse(value()?)?),
        3 => Transform::Blue(super::model::Percentage::parse(value()?)?),
        4 => Transform::BlueMod(super::model::Percentage::parse(value()?)?),
        5 => Transform::BlueOff(super::model::Percentage::parse(value()?)?),
        6 => Transform::Complement,
        7 => Transform::Gamma,
        8 => Transform::Gray,
        9 => Transform::Green(super::model::Percentage::parse(value()?)?),
        10 => Transform::GreenMod(super::model::Percentage::parse(value()?)?),
        11 => Transform::GreenOff(super::model::Percentage::parse(value()?)?),
        12 => Transform::Hue(PositiveAngle::parse(value()?)?),
        13 => Transform::HueMod(PositivePercentage::parse(value()?)?),
        14 => Transform::HueOff(Angle::parse(value()?)?),
        15 => Transform::Inverse,
        16 => Transform::InverseGamma,
        17 => Transform::Lum(super::model::Percentage::parse(value()?)?),
        18 => Transform::LumMod(super::model::Percentage::parse(value()?)?),
        19 => Transform::LumOff(super::model::Percentage::parse(value()?)?),
        20 => Transform::Red(super::model::Percentage::parse(value()?)?),
        21 => Transform::RedMod(super::model::Percentage::parse(value()?)?),
        22 => Transform::RedOff(super::model::Percentage::parse(value()?)?),
        23 => Transform::Sat(super::model::Percentage::parse(value()?)?),
        24 => Transform::SatMod(super::model::Percentage::parse(value()?)?),
        25 => Transform::SatOff(super::model::Percentage::parse(value()?)?),
        26 => Transform::Shade(PositiveFixedPercentage::parse(value()?)?),
        27 => Transform::Tint(PositiveFixedPercentage::parse(value()?)?),
        _ => unreachable!("known color transform kind is closed"),
    }))
}

fn push_transform(
    transforms: &mut Vec<Transform>,
    transform: Transform,
    max_transforms: usize,
) -> Result<()> {
    if transforms.len() >= max_transforms {
        return Err(Error::limit(
            crate::LimitKind::ColorTransforms,
            transforms.len() + 1,
            max_transforms,
        ));
    }
    transforms.push(transform);
    Ok(())
}

fn attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    allowed: &[&[u8]],
) -> Result<Option<Vec<(Vec<u8>, String)>>> {
    let mut values: Vec<(Vec<u8>, String)> = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        if attribute.key.prefix().is_some() || !allowed.contains(&key) {
            return Ok(None);
        }
        if values.iter().any(|(name, _)| name.as_slice() == key) {
            return Err(Error::Invalid(format!(
                "duplicate DrawingML color attribute '{}'",
                String::from_utf8_lossy(key)
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        values.push((key.to_vec(), value));
    }
    Ok(Some(values))
}

fn required<'a>(
    attributes: &'a [(Vec<u8>, String)],
    name: &[u8],
    element: &str,
) -> Result<&'a str> {
    optional(attributes, name).ok_or_else(|| {
        Error::Invalid(format!(
            "DrawingML {element} color is missing {}",
            String::from_utf8_lossy(name)
        ))
    })
}

fn optional<'a>(attributes: &'a [(Vec<u8>, String)], name: &[u8]) -> Option<&'a str> {
    attributes
        .iter()
        .find(|(key, _)| key.as_slice() == name)
        .map(|(_, value)| value.as_str())
}

fn validate_typed_limits(value: &Value, limits: Limits) -> Result<()> {
    let (nodes, depth, transforms) = match value {
        Value::Transformed(value) => (
            1usize + value.transforms().len(),
            2usize,
            value.transforms().len(),
        ),
        Value::Rgb(_)
        | Value::Scheme(_)
        | Value::ScRgb(_)
        | Value::Hsl(_)
        | Value::System(_)
        | Value::Preset(_)
        | Value::Unknown(_) => (1, 1, 0),
    };
    if transforms > limits.max_transforms {
        return Err(Error::limit(
            crate::LimitKind::ColorTransforms,
            transforms,
            limits.max_transforms,
        ));
    }
    if nodes > limits.max_nodes {
        return Err(Error::limit(
            crate::LimitKind::XmlNodes,
            nodes,
            limits.max_nodes,
        ));
    }
    if depth > limits.max_depth {
        return Err(Error::limit(
            crate::LimitKind::XmlDepth,
            depth,
            limits.max_depth,
        ));
    }
    Ok(())
}

fn write_value_xml<W: fmt::Write>(output: &mut W, value: &Value) -> fmt::Result {
    match value {
        Value::Rgb(rgb) => write_base(output, BaseRef::Rgb(*rgb), false),
        Value::Scheme(scheme) => write_base(output, BaseRef::Scheme(*scheme), false),
        Value::ScRgb(color) => write_base(output, BaseRef::ScRgb(*color), false),
        Value::Hsl(color) => write_base(output, BaseRef::Hsl(*color), false),
        Value::System(color) => write_base(output, BaseRef::System(color), false),
        Value::Preset(color) => write_base(output, BaseRef::Preset(color), false),
        Value::Transformed(value) => {
            write_base(output, BaseRef::from(value.base()), true)?;
            for transform in value.transforms() {
                write_transform(output, *transform)?;
            }
            write!(output, "</a:{}>", base_name(value.base()))
        },
        Value::Unknown(_) => unreachable!("opaque values are handled before typed serialization"),
    }
}

fn write_base<W: fmt::Write>(output: &mut W, base: BaseRef<'_>, transformed: bool) -> fmt::Result {
    match base {
        BaseRef::Rgb(rgb) => {
            write!(output, "<a:srgbClr val=\"{rgb}\"")?;
        },
        BaseRef::Scheme(scheme) => {
            write!(output, "<a:schemeClr val=\"{}\"", scheme.token())?;
        },
        BaseRef::ScRgb(color) => {
            write!(
                output,
                "<a:scrgbClr r=\"{}\" g=\"{}\" b=\"{}\"",
                color.red().value(),
                color.green().value(),
                color.blue().value()
            )?;
        },
        BaseRef::Hsl(color) => {
            write!(
                output,
                "<a:hslClr hue=\"{}\" sat=\"{}\" lum=\"{}\"",
                color.hue().value(),
                color.saturation().value(),
                color.luminance().value()
            )?;
        },
        BaseRef::System(color) => {
            write!(output, "<a:sysClr val=\"{}\"", color.token())?;
            if let Some(last) = color.last_rgb() {
                write!(output, " lastClr=\"{last}\"")?;
            }
        },
        BaseRef::Preset(color) => {
            write!(output, "<a:prstClr val=\"{}\"", color.token())?;
        },
    }
    if transformed {
        output.write_char('>')?;
    } else {
        output.write_str("/>")?;
    }
    Ok(())
}

fn base_name(base: &Base) -> &'static str {
    match base {
        Base::Rgb(_) => "srgbClr",
        Base::Scheme(_) => "schemeClr",
        Base::ScRgb(_) => "scrgbClr",
        Base::Hsl(_) => "hslClr",
        Base::System(_) => "sysClr",
        Base::Preset(_) => "prstClr",
    }
}

fn write_transform<W: fmt::Write>(output: &mut W, transform: Transform) -> fmt::Result {
    match transform {
        Transform::Alpha(value) => write_value(output, "alpha", value.value()),
        Transform::AlphaMod(value) => write_value(output, "alphaMod", value.value()),
        Transform::AlphaOff(value) => write_value(output, "alphaOff", value.value()),
        Transform::Blue(value) => write_value(output, "blue", value.value()),
        Transform::BlueMod(value) => write_value(output, "blueMod", value.value()),
        Transform::BlueOff(value) => write_value(output, "blueOff", value.value()),
        Transform::Complement => write_empty(output, "comp"),
        Transform::Gamma => write_empty(output, "gamma"),
        Transform::Gray => write_empty(output, "gray"),
        Transform::Green(value) => write_value(output, "green", value.value()),
        Transform::GreenMod(value) => write_value(output, "greenMod", value.value()),
        Transform::GreenOff(value) => write_value(output, "greenOff", value.value()),
        Transform::Hue(value) => write_value(output, "hue", value.value()),
        Transform::HueMod(value) => write_value(output, "hueMod", value.value()),
        Transform::HueOff(value) => write_value(output, "hueOff", value.value()),
        Transform::Inverse => write_empty(output, "inv"),
        Transform::InverseGamma => write_empty(output, "invGamma"),
        Transform::Lum(value) => write_value(output, "lum", value.value()),
        Transform::LumMod(value) => write_value(output, "lumMod", value.value()),
        Transform::LumOff(value) => write_value(output, "lumOff", value.value()),
        Transform::Red(value) => write_value(output, "red", value.value()),
        Transform::RedMod(value) => write_value(output, "redMod", value.value()),
        Transform::RedOff(value) => write_value(output, "redOff", value.value()),
        Transform::Sat(value) => write_value(output, "sat", value.value()),
        Transform::SatMod(value) => write_value(output, "satMod", value.value()),
        Transform::SatOff(value) => write_value(output, "satOff", value.value()),
        Transform::Shade(value) => write_value(output, "shade", value.value()),
        Transform::Tint(value) => write_value(output, "tint", value.value()),
    }
}

fn write_value<W: fmt::Write, V: fmt::Display>(
    output: &mut W,
    name: &str,
    value: V,
) -> fmt::Result {
    write!(output, "<a:{name} val=\"{value}\"/>")
}

fn write_empty<W: fmt::Write>(output: &mut W, name: &str) -> fmt::Result {
    write!(output, "<a:{name}/>")
}

fn format_error(error: fmt::Error) -> Error {
    Error::Xml(error.to_string())
}

#[derive(Default)]
struct LengthWriter {
    len: usize,
    overflow: bool,
}

impl fmt::Write for LengthWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        match self.len.checked_add(value.len()) {
            Some(length) => self.len = length,
            None => {
                self.len = usize::MAX;
                self.overflow = true;
            },
        }
        Ok(())
    }
}

fn tail_is_empty(reader: &mut Reader<&[u8]>) -> Result<bool> {
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Text(text)
                if text
                    .decode()
                    .map_err(|error| xml_error(&error))?
                    .trim()
                    .is_empty() => {},
            Event::Eof => return Ok(true),
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => return Ok(false),
        }
    }
}

fn xml_error(error: &quick_xml::encoding::EncodingError) -> Error {
    Error::Xml(error.to_string())
}
