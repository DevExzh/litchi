//! Bounded theme XML codecs.

use std::ops::Range;

use litchi_ooxml_common::mce::process_ooxml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::{Error, Result};

use super::model::{
    Color, Face, FontSet, Override, Palette, Part, Slot, System, validate_fonts, validate_name,
    validate_palette,
};

pub const NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub const STRICT_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;

const FORMAT_SCHEME: &str = "<a:fmtScheme name=\"Office\"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>";

/// Encode a complete theme part.
pub fn encode_part(name: &str, colors: &Palette, fonts: &FontSet) -> Result<Vec<u8>> {
    validate_name("theme", name)?;
    validate_palette(colors)?;
    validate_fonts(fonts)?;
    let mut xml = String::with_capacity(4096);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    xml.push_str("<a:theme xmlns:a=\"");
    xml.push_str(NAMESPACE);
    xml.push_str("\" name=\"");
    escape(&mut xml, name);
    xml.push_str("\"><a:themeElements>");
    push_palette(&mut xml, colors, false);
    push_fonts(&mut xml, fonts, false);
    xml.push_str(FORMAT_SCHEME);
    xml.push_str("</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>");
    bounded(xml.into_bytes(), "theme XML")
}

/// Encode a theme override part.
pub fn encode_override(value: &Override) -> Result<Vec<u8>> {
    if value.colors.is_none() && value.fonts.is_none() {
        return Err(invalid("theme override requires at least one scheme"));
    }
    if let Some(colors) = &value.colors {
        validate_palette(colors)?;
    }
    if let Some(fonts) = &value.fonts {
        validate_fonts(fonts)?;
    }
    let mut xml = String::with_capacity(2048);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    xml.push_str("<a:themeOverride xmlns:a=\"");
    xml.push_str(NAMESPACE);
    xml.push_str("\">");
    if let Some(colors) = &value.colors {
        push_palette(&mut xml, colors, false);
    }
    if let Some(fonts) = &value.fonts {
        push_fonts(&mut xml, fonts, false);
    }
    xml.push_str("</a:themeOverride>");
    bounded(xml.into_bytes(), "theme override XML")
}

/// Parse a complete theme part.
pub fn read(xml: &[u8]) -> Result<Part> {
    let parsed = parse(xml, "theme")?;
    let colors = parsed
        .colors
        .ok_or_else(|| invalid("theme has no color palette"))?;
    let fonts = parsed
        .fonts
        .ok_or_else(|| invalid("theme has no font set"))?;
    Ok(Part {
        name: parsed.name.unwrap_or_default(),
        colors,
        fonts,
    })
}

/// Parse a theme override, retaining only its typed color and font schemes.
pub fn read_override(xml: &[u8]) -> Result<Override> {
    let parsed = parse(xml, "themeOverride")?;
    Ok(Override {
        colors: parsed.colors,
        fonts: parsed.fonts,
    })
}

/// Replace a direct `clrScheme` or `fontScheme` child of `themeElements`.
pub fn replace_scheme(xml: &[u8], local: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let range = direct_scheme_range(xml, local)?
        .ok_or_else(|| invalid("theme scheme is missing from themeElements"))?;
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[range.end..]);
    bounded(output, "patched theme XML")
}

pub fn encode_palette_fragment(value: &Palette) -> Result<Vec<u8>> {
    validate_palette(value)?;
    let mut output = String::with_capacity(1024);
    push_palette(&mut output, value, true);
    bounded(output.into_bytes(), "color palette XML")
}

pub fn encode_fonts_fragment(value: &FontSet) -> Result<Vec<u8>> {
    validate_fonts(value)?;
    let mut output = String::with_capacity(1024);
    push_fonts(&mut output, value, true);
    bounded(output.into_bytes(), "font set XML")
}

fn push_palette(xml: &mut String, value: &Palette, declare_namespace: bool) {
    xml.push_str("<a:clrScheme");
    if declare_namespace {
        xml.push_str(" xmlns:a=\"");
        xml.push_str(NAMESPACE);
        xml.push('"');
    }
    xml.push_str(" name=\"");
    escape(xml, value.name());
    xml.push_str("\">");
    for slot in Slot::ALL {
        xml.push_str("<a:");
        xml.push_str(slot.token());
        xml.push('>');
        match value
            .color(slot)
            .expect("validated palettes contain all slots")
        {
            Color::Rgb(value) => {
                xml.push_str("<a:srgbClr val=\"");
                xml.push_str(value);
                xml.push_str("\"/>");
            },
            Color::System { kind, last } => {
                xml.push_str("<a:sysClr val=\"");
                xml.push_str(kind.token());
                if let Some(last) = last {
                    xml.push_str("\" lastClr=\"");
                    xml.push_str(last);
                }
                xml.push_str("\"/>");
            },
        }
        xml.push_str("</a:");
        xml.push_str(slot.token());
        xml.push('>');
    }
    xml.push_str("</a:clrScheme>");
}

fn push_fonts(xml: &mut String, value: &FontSet, declare_namespace: bool) {
    xml.push_str("<a:fontScheme");
    if declare_namespace {
        xml.push_str(" xmlns:a=\"");
        xml.push_str(NAMESPACE);
        xml.push('"');
    }
    xml.push_str(" name=\"");
    escape(xml, value.name());
    xml.push_str("\">");
    push_face(xml, "majorFont", value.major());
    push_face(xml, "minorFont", value.minor());
    xml.push_str("</a:fontScheme>");
}

fn push_face(xml: &mut String, name: &str, face: &Face) {
    xml.push_str("<a:");
    xml.push_str(name);
    xml.push_str("><a:latin typeface=\"");
    escape(xml, &face.latin);
    xml.push_str("\"/><a:ea typeface=\"");
    escape(xml, &face.east_asian);
    xml.push_str("\"/><a:cs typeface=\"");
    escape(xml, &face.complex_script);
    xml.push_str("\"/>");
    for script in &face.scripts {
        xml.push_str("<a:font script=\"");
        escape(xml, &script.code);
        xml.push_str("\" typeface=\"");
        escape(xml, &script.typeface);
        xml.push_str("\"/>");
    }
    xml.push_str("</a:");
    xml.push_str(name);
    xml.push('>');
}

#[derive(Default)]
struct Parsed {
    name: Option<String>,
    colors: Option<Palette>,
    fonts: Option<FontSet>,
}

struct FontState {
    name: String,
    major: Option<Face>,
    minor: Option<Face>,
    target: Option<bool>,
}

fn parse(xml: &[u8], root_name: &str) -> Result<Parsed> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_XML_BYTES {
        return Err(limit("processed theme XML bytes", MAX_XML_BYTES));
    }
    let mut reader = Reader::from_reader(processed.as_ref());
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut parsed = Parsed::default();
    let mut palette_name = String::new();
    let mut palette_values: Vec<(Slot, Color)> = Vec::new();
    let mut current_slot = None;
    let mut fonts: Option<FontState> = None;
    let mut nodes = 0usize;
    let scheme_depth = usize::from(root_name == "theme") + 1;
    let mut root_seen = false;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme node count overflow"))?;
                if nodes > MAX_NODES {
                    return Err(limit("theme XML nodes", MAX_NODES));
                }
                let local = element.local_name().as_ref().to_vec();
                let depth = stack.len();
                if !root_seen {
                    if local != root_name.as_bytes() {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                    parsed.name = attr(&element, b"name", reader.decoder())?;
                } else {
                    open(
                        &local,
                        depth,
                        scheme_depth,
                        &element,
                        reader.decoder(),
                        &mut palette_name,
                        &mut palette_values,
                        &mut current_slot,
                        &mut fonts,
                        &mut parsed,
                    )?;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("theme XML depth", MAX_DEPTH));
                }
                stack.push(local);
            },
            Event::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("theme node count overflow"))?;
                if nodes > MAX_NODES {
                    return Err(limit("theme XML nodes", MAX_NODES));
                }
                let local = element.local_name().as_ref().to_vec();
                let depth = stack.len();
                if !root_seen {
                    if local != root_name.as_bytes() {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                    parsed.name = attr(&element, b"name", reader.decoder())?;
                } else {
                    open(
                        &local,
                        depth,
                        scheme_depth,
                        &element,
                        reader.decoder(),
                        &mut palette_name,
                        &mut palette_values,
                        &mut current_slot,
                        &mut fonts,
                        &mut parsed,
                    )?;
                    close(
                        &local,
                        depth,
                        scheme_depth,
                        &mut palette_name,
                        &mut palette_values,
                        &mut current_slot,
                        &mut fonts,
                        &mut parsed,
                    )?;
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                let Some(open_name) = stack.pop() else {
                    return Err(invalid("theme XML has an unexpected closing element"));
                };
                if open_name.as_slice() != local.as_ref() {
                    return Err(invalid("theme XML closing element does not match"));
                }
                let depth = stack.len();
                close(
                    local.as_ref(),
                    depth,
                    scheme_depth,
                    &mut palette_name,
                    &mut palette_values,
                    &mut current_slot,
                    &mut fonts,
                    &mut parsed,
                )?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("theme XML contains forbidden markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !stack.is_empty() {
        return Err(invalid("theme XML is unterminated"));
    }
    Ok(parsed)
}

#[allow(clippy::too_many_arguments)]
fn open(
    local: &[u8],
    depth: usize,
    scheme_depth: usize,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    palette_name: &mut String,
    palette_values: &mut Vec<(Slot, Color)>,
    current_slot: &mut Option<Slot>,
    fonts: &mut Option<FontState>,
    parsed: &mut Parsed,
) -> Result<()> {
    if depth == scheme_depth && local == b"clrScheme" {
        if parsed.colors.is_some() {
            return Err(invalid("theme contains multiple color palettes"));
        }
        palette_name.clear();
        palette_name.push_str(
            attr(element, b"name", decoder)?
                .as_deref()
                .unwrap_or_default(),
        );
    } else if depth == scheme_depth && local == b"fontScheme" {
        if fonts.is_some() || parsed.fonts.is_some() {
            return Err(invalid("theme contains multiple font sets"));
        }
        *fonts = Some(FontState {
            name: attr(element, b"name", decoder)?.unwrap_or_default(),
            major: None,
            minor: None,
            target: None,
        });
    } else if parsed.colors.is_none() && depth == scheme_depth + 1 {
        if let Some(slot) = std::str::from_utf8(local).ok().and_then(Slot::from_token) {
            if current_slot.replace(slot).is_some() {
                return Err(invalid("theme has nested color slots"));
            }
        }
    } else if parsed.colors.is_none() && depth == scheme_depth + 2 && current_slot.is_some() {
        let color = match local {
            b"srgbClr" => Color::rgb(
                &attr(element, b"val", decoder)?.ok_or_else(|| invalid("srgbClr lacks val"))?,
            )?,
            b"sysClr" => {
                let kind = attr(element, b"val", decoder)?
                    .and_then(|value| System::from_token(&value))
                    .ok_or_else(|| invalid("sysClr has an unknown system color"))?;
                Color::system(kind, attr(element, b"lastClr", decoder)?.as_deref())?
            },
            _ => return Ok(()),
        };
        let slot = current_slot.expect("checked above");
        if palette_values.iter().any(|(existing, _)| *existing == slot) {
            return Err(invalid("theme contains a duplicate color slot"));
        }
        palette_values.push((slot, color));
    } else if let Some(fonts) = fonts.as_mut() {
        if depth == scheme_depth + 1 && local == b"majorFont" {
            if fonts.major.is_some() {
                return Err(invalid("font set has multiple major faces"));
            }
            fonts.major = Some(Face::new(""));
            fonts.target = Some(false);
        } else if depth == scheme_depth + 1 && local == b"minorFont" {
            if fonts.minor.is_some() {
                return Err(invalid("font set has multiple minor faces"));
            }
            fonts.minor = Some(Face::new(""));
            fonts.target = Some(true);
        } else if depth == scheme_depth + 2 {
            let Some(target) = fonts.target else {
                return Ok(());
            };
            let face = if target {
                fonts.minor.as_mut()
            } else {
                fonts.major.as_mut()
            }
            .ok_or_else(|| invalid("font face is missing"))?;
            match local {
                b"latin" => {
                    face.latin = attr(element, b"typeface", decoder)?
                        .ok_or_else(|| invalid("latin face lacks typeface"))?
                },
                b"ea" => face.east_asian = attr(element, b"typeface", decoder)?.unwrap_or_default(),
                b"cs" => {
                    face.complex_script = attr(element, b"typeface", decoder)?.unwrap_or_default()
                },
                b"font" => face.scripts.push(super::model::Script {
                    code: attr(element, b"script", decoder)?.unwrap_or_default(),
                    typeface: attr(element, b"typeface", decoder)?
                        .ok_or_else(|| invalid("script face lacks typeface"))?,
                }),
                _ => {},
            }
        }
    }
    Ok(())
}

fn close(
    local: &[u8],
    depth: usize,
    scheme_depth: usize,
    palette_name: &mut String,
    palette_values: &mut Vec<(Slot, Color)>,
    current_slot: &mut Option<Slot>,
    fonts: &mut Option<FontState>,
    parsed: &mut Parsed,
) -> Result<()> {
    if local == b"clrScheme" && depth == scheme_depth {
        let palette = Palette::new(std::mem::take(palette_name));
        let palette = palette_values
            .drain(..)
            .fold(palette, |palette, (slot, color)| palette.with(slot, color));
        validate_palette(&palette)?;
        parsed.colors = Some(palette);
    } else if local == b"fontScheme" && depth == scheme_depth {
        let fonts = fonts
            .take()
            .ok_or_else(|| invalid("font set parser state is missing"))?;
        let major = fonts
            .major
            .ok_or_else(|| invalid("font set lacks a major face"))?;
        let minor = fonts
            .minor
            .ok_or_else(|| invalid("font set lacks a minor face"))?;
        let value = FontSet::new(fonts.name, major, minor);
        validate_fonts(&value)?;
        parsed.fonts = Some(value);
    } else if current_slot.is_some()
        && Slot::from_token(std::str::from_utf8(local).unwrap_or_default()).is_some()
    {
        *current_slot = None;
    } else if let Some(fonts) = fonts.as_mut()
        && matches!(local, b"majorFont" | b"minorFont")
        && depth == scheme_depth + 1
    {
        fonts.target = None;
    }
    Ok(())
}

fn attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    Ok(unqualified_attribute_value(element, name, decoder)?)
}

fn direct_scheme_range(xml: &[u8], target: &[u8]) -> Result<Option<Range<usize>>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<(usize, Vec<u8>)> = Vec::new();
    let mut nodes = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("theme XML offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("theme XML offset exceeds usize"))?;
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| invalid("theme node count overflow"))?;
        if nodes > MAX_NODES {
            return Err(limit("theme XML nodes", MAX_NODES));
        }
        match event {
            Event::Start(element) => {
                let local = element.local_name().as_ref().to_vec();
                if stack.len() == 2 && local.as_slice() == target {
                    stack.push((start, local));
                } else {
                    stack.push((start, local));
                }
            },
            Event::Empty(element)
                if stack.len() == 2 && element.local_name().as_ref() == target =>
            {
                return Ok(Some(start..end));
            },
            Event::Empty(_) => {},
            Event::End(element) => {
                let Some((open, local)) = stack.pop() else {
                    return Err(invalid("theme XML nesting underflow"));
                };
                if local.as_slice() != element.local_name().as_ref() {
                    return Err(invalid("theme XML closing element does not match"));
                }
                if stack.len() == 2 && local.as_slice() == target {
                    return Ok(Some(open..end));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("theme XML contains forbidden markup"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("theme XML is unterminated"));
    }
    Ok(None)
}

fn bounded(value: Vec<u8>, resource: &'static str) -> Result<Vec<u8>> {
    if value.len() > MAX_XML_BYTES {
        return Err(limit(resource, MAX_XML_BYTES));
    }
    Ok(value)
}

fn escape(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            character => output.push(character),
        }
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shape::theme::model::{Color, Face, FontSet, Palette, Slot, System};

    fn palette() -> Palette {
        Slot::ALL
            .into_iter()
            .fold(Palette::new("Office"), |palette, slot| {
                let color = if slot == Slot::Dark1 {
                    Color::system(System::WindowText, Some("000000")).unwrap()
                } else {
                    Color::rgb("4F81BD").unwrap()
                };
                palette.with(slot, color)
            })
    }

    #[test]
    fn complete_theme_round_trips() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let xml = encode_part("Office Theme", &palette(), &fonts).unwrap();
        let parsed = read(&xml).unwrap();
        assert_eq!(parsed.name, "Office Theme");
        assert_eq!(parsed.colors, palette());
        assert_eq!(parsed.fonts, fonts);
    }

    #[test]
    fn override_round_trips_and_scheme_patching_is_local() {
        let colors = palette();
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let override_value = Override::new().colors(colors.clone()).fonts(fonts.clone());
        assert_eq!(
            read_override(&encode_override(&override_value).unwrap()).unwrap(),
            override_value
        );
        let original = encode_part("Office", &colors, &fonts).unwrap();
        let changed = Palette::new("Changed");
        let changed = Slot::ALL.into_iter().fold(changed, |palette, slot| {
            palette.with(slot, Color::rgb("FFFFFF").unwrap())
        });
        let fragment = encode_palette_fragment(&changed).unwrap();
        let patched = replace_scheme(&original, b"clrScheme", &fragment).unwrap();
        assert_eq!(read(&patched).unwrap().colors, changed);
        assert!(
            patched
                .windows(FORMAT_SCHEME.len())
                .any(|window| window == FORMAT_SCHEME.as_bytes())
        );
    }
}
