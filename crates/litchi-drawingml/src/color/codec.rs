//! Bounded XML codec for DrawingML color fragments.

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::{Error, Result};

use super::model::{
    Angle, Base, FixedPercentage, Hsl, PositiveAngle, PositiveFixedPercentage, PositivePercentage,
    Preset, Rgb, ScRgb, Scheme, System, Transform, Transformed, Unknown, Value, write_hex,
};
use super::validation;

pub use super::validation::{MAX_DEPTH, MAX_NODES, MAX_TRANSFORMS, MAX_XML_BYTES};

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

/// Read one DrawingML color-choice fragment.
pub fn read(xml: &[u8]) -> Result<Value> {
    let validated = validation::validated_fragment(xml)?;
    let mut reader = Reader::from_reader(validated);

    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::Start(element) => {
                return Ok(read_choice(&mut reader, &element)?
                    .unwrap_or_else(|| Value::Unknown(Unknown::from_validated(validated))));
            },
            Event::Empty(element) => {
                return Ok(parse_empty_choice(&mut reader, &element)?
                    .unwrap_or_else(|| Value::Unknown(Unknown::from_validated(validated))));
            },
            _ => return Ok(Value::Unknown(Unknown::from_validated(validated))),
        }
    }
}

/// Write one DrawingML color-choice fragment using the conventional `a`
/// prefix. The result is a fragment and intentionally has no namespace
/// declaration so the host can retain its own namespace spelling.
pub fn write(value: &Value) -> Result<Vec<u8>> {
    let mut output = String::new();
    match value {
        Value::Rgb(rgb) => write_base(&mut output, BaseRef::Rgb(*rgb), false),
        Value::Scheme(scheme) => write_base(&mut output, BaseRef::Scheme(*scheme), false),
        Value::ScRgb(color) => write_base(&mut output, BaseRef::ScRgb(*color), false),
        Value::Hsl(color) => write_base(&mut output, BaseRef::Hsl(*color), false),
        Value::System(color) => write_base(&mut output, BaseRef::System(color), false),
        Value::Preset(color) => write_base(&mut output, BaseRef::Preset(color), false),
        Value::Transformed(value) => {
            write_base(&mut output, BaseRef::from(value.base()), true);
            for transform in value.transforms() {
                write_transform(&mut output, *transform);
            }
            output.push_str("</a:");
            output.push_str(base_name(value.base()));
            output.push('>');
        },
        Value::Unknown(value) => return Ok(value.as_xml().to_vec()),
    }

    if output.len() > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "DrawingML color XML",
            limit: MAX_XML_BYTES,
        });
    }
    Ok(output.into_bytes())
}

fn read_choice(reader: &mut Reader<&[u8]>, element: &BytesStart<'_>) -> Result<Option<Value>> {
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
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::Empty(element) => {
                let Some(transform) = parse_transform(&element, reader.decoder())? else {
                    return Ok(None);
                };
                push_transform(&mut transforms, transform)?;
            },
            Event::Start(element) => {
                let Some(transform) = parse_started_transform(reader, &element)? else {
                    return Ok(None);
                };
                push_transform(&mut transforms, transform)?;
            },
            Event::End(element) if element.name().as_ref() == root_name.as_slice() => break,
            Event::Eof => {
                return Err(Error::Invalid(
                    "DrawingML color choice ended before its root element".into(),
                ));
            },
            _ => return Ok(None),
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
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::End(end) if end.name().as_ref() == name.as_slice() => {
                return Ok(Some(transform));
            },
            _ => return Ok(None),
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

fn push_transform(transforms: &mut Vec<Transform>, transform: Transform) -> Result<()> {
    if transforms.len() >= MAX_TRANSFORMS {
        return Err(Error::Limit {
            resource: "DrawingML color transforms",
            limit: MAX_TRANSFORMS,
        });
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
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        if attribute.key.prefix().is_some() || !allowed.iter().any(|expected| *expected == key) {
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

fn write_base(output: &mut String, base: BaseRef<'_>, transformed: bool) {
    match base {
        BaseRef::Rgb(rgb) => {
            output.push_str("<a:srgbClr val=\"");
            write_hex(output, rgb);
            output.push('"');
        },
        BaseRef::Scheme(scheme) => {
            output.push_str("<a:schemeClr val=\"");
            output.push_str(scheme.token());
            output.push('"');
        },
        BaseRef::ScRgb(color) => {
            output.push_str("<a:scrgbClr r=\"");
            output.push_str(&color.red().value().to_string());
            output.push_str("\" g=\"");
            output.push_str(&color.green().value().to_string());
            output.push_str("\" b=\"");
            output.push_str(&color.blue().value().to_string());
            output.push('"');
        },
        BaseRef::Hsl(color) => {
            output.push_str("<a:hslClr hue=\"");
            output.push_str(&color.hue().value().to_string());
            output.push_str("\" sat=\"");
            output.push_str(&color.saturation().value().to_string());
            output.push_str("\" lum=\"");
            output.push_str(&color.luminance().value().to_string());
            output.push('"');
        },
        BaseRef::System(color) => {
            output.push_str("<a:sysClr val=\"");
            output.push_str(color.token());
            output.push('"');
            if let Some(last) = color.last_rgb() {
                output.push_str(" lastClr=\"");
                write_hex(output, last);
                output.push('"');
            }
        },
        BaseRef::Preset(color) => {
            output.push_str("<a:prstClr val=\"");
            output.push_str(color.token());
            output.push('"');
        },
    }
    if transformed {
        output.push('>');
    } else {
        output.push_str("/>");
    }
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

fn write_transform(output: &mut String, transform: Transform) {
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

fn write_value(output: &mut String, name: &str, value: impl ToString) {
    output.push_str("<a:");
    output.push_str(name);
    output.push_str(" val=\"");
    output.push_str(&value.to_string());
    output.push_str("\"/>");
}

fn write_empty(output: &mut String, name: &str) {
    output.push_str("<a:");
    output.push_str(name);
    output.push_str("/>");
}

fn tail_is_empty(reader: &mut Reader<&[u8]>) -> Result<bool> {
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Text(text) if text.decode().map_err(xml_error)?.trim().is_empty() => {},
            Event::Eof => return Ok(true),
            _ => return Ok(false),
        }
    }
}

fn xml_error(error: quick_xml::encoding::EncodingError) -> Error {
    Error::Xml(error.to_string())
}
