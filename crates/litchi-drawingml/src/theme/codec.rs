//! Bounded `DrawingML` theme XML codecs.

use std::ops::Range;

use litchi_ooxml_common::mce::process_ooxml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::{NsReader, Reader};

use crate::{Error, Result};

use super::model::{
    Color, Face, FontSet, Override, Palette, Slot, System, Theme, validate_fonts, validate_name,
    validate_palette,
};

pub const NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub const STRICT_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;
const MAX_NAMESPACE_DECLARATIONS: usize = 256;

// `CT_StyleMatrix` requires at least three entries in every style list. Keep
// the authored default deliberately small while still valid in both the
// Transitional and Strict DrawingML schemas.
const FORMAT_SCHEME: &str = concat!(
    "<a:fmtScheme name=\"Office\">",
    "<a:fillStyleLst>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "</a:fillStyleLst>",
    "<a:lnStyleLst>",
    "<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:prstDash val=\"solid\"/><a:miter lim=\"80%\"/></a:ln>",
    "<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:prstDash val=\"solid\"/><a:miter lim=\"80%\"/></a:ln>",
    "<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:prstDash val=\"solid\"/><a:miter lim=\"80%\"/></a:ln>",
    "</a:lnStyleLst>",
    "<a:effectStyleLst>",
    "<a:effectStyle><a:effectLst/></a:effectStyle>",
    "<a:effectStyle><a:effectLst/></a:effectStyle>",
    "<a:effectStyle><a:effectLst/></a:effectStyle>",
    "</a:effectStyleLst>",
    "<a:bgFillStyleLst>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
    "</a:bgFillStyleLst>",
    "</a:fmtScheme>"
);

/// Encode a complete theme part.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn read(xml: &[u8]) -> Result<Theme> {
    let parsed = parse(xml, "theme")?;
    let colors = parsed
        .colors
        .ok_or_else(|| invalid("theme has no color palette"))?;
    let fonts = parsed
        .fonts
        .ok_or_else(|| invalid("theme has no font set"))?;
    Ok(Theme {
        name: parsed.name.unwrap_or_default(),
        colors,
        fonts,
    })
}

/// Parse a theme override, retaining only its typed color and font schemes.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn read_override(xml: &[u8]) -> Result<Override> {
    let parsed = parse(xml, "themeOverride")?;
    Ok(Override {
        colors: parsed.colors,
        fonts: parsed.fonts,
    })
}

/// Replace a direct `clrScheme` or `fontScheme` child of `themeElements`.
///
/// Replacement is refused when the source scheme contains namespace-qualified,
/// unknown, or otherwise unmodeled attributes or children. This keeps a typed
/// replacement from silently dropping source data that the shared model cannot
/// represent.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn replace_scheme(xml: &[u8], local: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let range = scheme_replacement_range(xml, local)?;
    let source_namespace = theme_root_namespace(xml)?;
    let adjusted_replacement = adjust_replacement_namespace(replacement, source_namespace)?;
    let replacement = adjusted_replacement.as_deref().unwrap_or(replacement);
    let output_len = range
        .start
        .checked_add(replacement.len())
        .and_then(|length| {
            xml.len()
                .checked_sub(range.end)
                .and_then(|suffix| length.checked_add(suffix))
        })
        .ok_or_else(|| invalid("patched theme XML length overflows"))?;
    if output_len > MAX_XML_BYTES {
        return Err(limit("patched theme XML", MAX_XML_BYTES));
    }
    let mut output = Vec::with_capacity(output_len);
    output.extend_from_slice(&xml[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[range.end..]);
    bounded(output, "patched theme XML")
}

/// Find one directly owned, losslessly replaceable scheme in a complete Theme
/// part without allocating a patched copy.
///
/// The returned byte range covers the complete source element. Namespace,
/// parent, duplicate-target, depth, size, and unsupported-content checks are
/// identical to [`replace_scheme`], so callers can stage several ranges before
/// publishing one checked output.
///
/// # Errors
///
/// Returns an error when the source is malformed, the target is absent or
/// ambiguous, or replacing it would discard source data outside the typed
/// theme model.
pub fn scheme_replacement_range(xml: &[u8], local: &[u8]) -> Result<Range<usize>> {
    direct_scheme_range(xml, local)?
        .ok_or_else(|| invalid("theme scheme is missing from themeElements"))
}

/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn encode_palette_fragment(value: &Palette) -> Result<Vec<u8>> {
    validate_palette(value)?;
    let mut output = String::with_capacity(1024);
    push_palette(&mut output, value, true);
    bounded(output.into_bytes(), "color palette XML")
}

/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
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
        let Some(color) = value.color(slot) else {
            continue;
        };
        match color {
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

struct Frame {
    local: Vec<u8>,
    drawingml: bool,
}

struct FontState {
    name: String,
    major: Option<Face>,
    minor: Option<Face>,
    target: Option<bool>,
}

fn is_drawingml_namespace(namespace: &ResolveResult<'_>) -> bool {
    drawingml_namespace(namespace).is_some()
}

fn drawingml_namespace(namespace: &ResolveResult<'_>) -> Option<&'static [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == NAMESPACE.as_bytes() => {
            Some(NAMESPACE.as_bytes())
        },
        ResolveResult::Bound(Namespace(value)) if *value == STRICT_NAMESPACE.as_bytes() => {
            Some(STRICT_NAMESPACE.as_bytes())
        },
        _ => None,
    }
}

fn reject_unknown_namespace(namespace: &ResolveResult<'_>) -> Result<()> {
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid("theme XML uses an undeclared namespace prefix"));
    }
    Ok(())
}

fn is_parent(stack: &[Frame], local: &[u8]) -> bool {
    stack
        .last()
        .is_some_and(|frame| frame.drawingml && frame.local.as_slice() == local)
}

fn is_color_slot_parent(stack: &[Frame]) -> bool {
    stack.last().is_some_and(|frame| {
        frame.drawingml
            && std::str::from_utf8(&frame.local)
                .ok()
                .and_then(Slot::from_token)
                .is_some()
    })
}

fn is_font_face_parent(stack: &[Frame]) -> bool {
    stack.last().is_some_and(|frame| {
        frame.drawingml && matches!(frame.local.as_slice(), b"majorFont" | b"minorFont")
    })
}

fn parse(xml: &[u8], root_name: &str) -> Result<Parsed> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_XML_BYTES {
        return Err(limit("processed theme XML bytes", MAX_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().check_end_names = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    let mut stack: Vec<Frame> = Vec::new();
    let mut parsed = Parsed::default();
    let mut palette_name = String::new();
    let mut palette_values: Vec<(Slot, Color)> = Vec::new();
    let mut current_slot = None;
    let mut fonts: Option<FontState> = None;
    let mut nodes = 0usize;
    let scheme_depth = usize::from(root_name == "theme") + 1;
    let scheme_parent = if root_name == "theme" {
        b"themeElements".as_slice()
    } else {
        root_name.as_bytes()
    };
    let mut root_seen = false;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        reject_unknown_namespace(&namespace)?;
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
                let drawingml = is_drawingml_namespace(&namespace);
                if root_seen {
                    if stack.is_empty() {
                        return Err(invalid("theme XML has multiple root elements"));
                    }
                    open(
                        &local,
                        drawingml,
                        &stack,
                        depth,
                        scheme_depth,
                        scheme_parent,
                        &element,
                        reader.decoder(),
                        &mut palette_name,
                        &mut palette_values,
                        &mut current_slot,
                        &mut fonts,
                        &mut parsed,
                    )?;
                } else {
                    if local != root_name.as_bytes() || !drawingml {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                    parsed.name = attr(&element, b"name", reader.decoder())?;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("theme XML depth", MAX_DEPTH));
                }
                stack.push(Frame { local, drawingml });
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
                let drawingml = is_drawingml_namespace(&namespace);
                if root_seen {
                    if stack.is_empty() {
                        return Err(invalid("theme XML has multiple root elements"));
                    }
                    open(
                        &local,
                        drawingml,
                        &stack,
                        depth,
                        scheme_depth,
                        scheme_parent,
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
                        drawingml,
                        &stack,
                        depth,
                        scheme_depth,
                        scheme_parent,
                        &mut palette_name,
                        &mut palette_values,
                        &mut current_slot,
                        &mut fonts,
                        &mut parsed,
                    )?;
                } else {
                    if local != root_name.as_bytes() || !drawingml {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                    parsed.name = attr(&element, b"name", reader.decoder())?;
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                let drawingml = is_drawingml_namespace(&namespace);
                let Some(frame) = stack.pop() else {
                    return Err(invalid("theme XML has an unexpected closing element"));
                };
                if frame.local.as_slice() != local.as_ref() || frame.drawingml != drawingml {
                    return Err(invalid("theme XML closing element does not match"));
                }
                let depth = stack.len();
                close(
                    local.as_ref(),
                    drawingml,
                    &stack,
                    depth,
                    scheme_depth,
                    scheme_parent,
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
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen || !stack.is_empty() {
        return Err(invalid("theme XML is unterminated"));
    }
    Ok(parsed)
}

#[allow(
    clippy::too_many_arguments,
    reason = "one theme start event updates the complete bounded schema parser state"
)]
fn open(
    local: &[u8],
    drawingml: bool,
    stack: &[Frame],
    depth: usize,
    scheme_depth: usize,
    scheme_parent: &[u8],
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    palette_name: &mut String,
    palette_values: &mut Vec<(Slot, Color)>,
    current_slot: &mut Option<Slot>,
    fonts: &mut Option<FontState>,
    parsed: &mut Parsed,
) -> Result<()> {
    if drawingml
        && depth == scheme_depth
        && local == b"clrScheme"
        && is_parent(stack, scheme_parent)
    {
        if parsed.colors.is_some() {
            return Err(invalid("theme contains multiple color palettes"));
        }
        palette_name.clear();
        palette_name.push_str(
            attr(element, b"name", decoder)?
                .as_deref()
                .unwrap_or_default(),
        );
    } else if drawingml
        && depth == scheme_depth
        && local == b"fontScheme"
        && is_parent(stack, scheme_parent)
    {
        if fonts.is_some() || parsed.fonts.is_some() {
            return Err(invalid("theme contains multiple font sets"));
        }
        *fonts = Some(FontState {
            name: attr(element, b"name", decoder)?.unwrap_or_default(),
            major: None,
            minor: None,
            target: None,
        });
    } else if drawingml
        && parsed.colors.is_none()
        && depth == scheme_depth + 1
        && is_parent(stack, b"clrScheme")
    {
        if let Some(slot) = std::str::from_utf8(local).ok().and_then(Slot::from_token)
            && current_slot.replace(slot).is_some()
        {
            return Err(invalid("theme has nested color slots"));
        }
    } else if drawingml
        && parsed.colors.is_none()
        && depth == scheme_depth + 2
        && current_slot.is_some()
        && is_color_slot_parent(stack)
    {
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
        let slot = current_slot.ok_or_else(|| invalid("theme color slot is missing"))?;
        if palette_values.iter().any(|(existing, _)| *existing == slot) {
            return Err(invalid("theme contains a duplicate color slot"));
        }
        palette_values.push((slot, color));
    } else if drawingml && let Some(fonts) = fonts.as_mut() {
        if depth == scheme_depth + 1 && local == b"majorFont" && is_parent(stack, b"fontScheme") {
            if fonts.major.is_some() {
                return Err(invalid("font set has multiple major faces"));
            }
            fonts.major = Some(Face::new(""));
            fonts.target = Some(false);
        } else if depth == scheme_depth + 1
            && local == b"minorFont"
            && is_parent(stack, b"fontScheme")
        {
            if fonts.minor.is_some() {
                return Err(invalid("font set has multiple minor faces"));
            }
            fonts.minor = Some(Face::new(""));
            fonts.target = Some(true);
        } else if depth == scheme_depth + 2 && is_font_face_parent(stack) {
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
                        .ok_or_else(|| invalid("latin face lacks typeface"))?;
                },
                b"ea" => face.east_asian = attr(element, b"typeface", decoder)?.unwrap_or_default(),
                b"cs" => {
                    face.complex_script = attr(element, b"typeface", decoder)?.unwrap_or_default();
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
    drawingml: bool,
    stack: &[Frame],
    depth: usize,
    scheme_depth: usize,
    scheme_parent: &[u8],
    palette_name: &mut String,
    palette_values: &mut Vec<(Slot, Color)>,
    current_slot: &mut Option<Slot>,
    fonts: &mut Option<FontState>,
    parsed: &mut Parsed,
) -> Result<()> {
    if drawingml
        && local == b"clrScheme"
        && depth == scheme_depth
        && is_parent(stack, scheme_parent)
    {
        let palette = Palette::new(std::mem::take(palette_name));
        let palette = palette_values
            .drain(..)
            .fold(palette, |palette, (slot, color)| palette.with(slot, color));
        validate_palette(&palette)?;
        parsed.colors = Some(palette);
    } else if drawingml
        && local == b"fontScheme"
        && depth == scheme_depth
        && is_parent(stack, scheme_parent)
    {
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
    } else if drawingml
        && current_slot.is_some()
        && depth == scheme_depth + 1
        && is_parent(stack, b"clrScheme")
        && Slot::from_token(std::str::from_utf8(local).unwrap_or_default()).is_some()
    {
        *current_slot = None;
    } else if drawingml
        && let Some(fonts) = fonts.as_mut()
        && matches!(local, b"majorFont" | b"minorFont")
        && depth == scheme_depth + 1
        && is_parent(stack, b"fontScheme")
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

fn theme_root_namespace(xml: &[u8]) -> Result<&'static [u8]> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != b"theme" {
                    return Err(invalid("theme XML has an unexpected root"));
                }
                return match namespace {
                    ResolveResult::Bound(Namespace(value)) if value == NAMESPACE.as_bytes() => {
                        Ok(NAMESPACE.as_bytes())
                    },
                    ResolveResult::Bound(Namespace(value))
                        if value == STRICT_NAMESPACE.as_bytes() =>
                    {
                        Ok(STRICT_NAMESPACE.as_bytes())
                    },
                    _ => Err(invalid("theme XML has an unsupported root namespace")),
                };
            },
            Event::DocType(_) | Event::PI(_) | Event::End(_) => {
                return Err(invalid("theme XML has forbidden markup before its root"));
            },
            Event::Eof => return Err(invalid("theme XML has no root")),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn adjust_replacement_namespace(
    replacement: &[u8],
    source_namespace: &[u8],
) -> Result<Option<Vec<u8>>> {
    let Some(range) = replacement_root_attribute_range(replacement, b"xmlns:a")? else {
        return Ok(None);
    };
    let current = &replacement[range.clone()];
    if current == source_namespace {
        return Ok(None);
    }
    if current != NAMESPACE.as_bytes() && current != STRICT_NAMESPACE.as_bytes() {
        return Err(invalid(
            "theme replacement has an unsupported DrawingML namespace",
        ));
    }
    let mut adjusted = Vec::with_capacity(
        replacement
            .len()
            .checked_sub(current.len())
            .and_then(|length| length.checked_add(source_namespace.len()))
            .ok_or_else(|| invalid("theme replacement length overflows"))?,
    );
    adjusted.extend_from_slice(&replacement[..range.start]);
    adjusted.extend_from_slice(source_namespace);
    adjusted.extend_from_slice(&replacement[range.end..]);
    Ok(Some(adjusted))
}

fn replacement_root_attribute_range(xml: &[u8], wanted: &[u8]) -> Result<Option<Range<usize>>> {
    let mut reader = Reader::from_reader(xml);
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("theme replacement offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("theme replacement offset exceeds usize"))?;
        match event {
            Event::Start(_) | Event::Empty(_) => {
                let raw = xml
                    .get(start..end)
                    .ok_or_else(|| invalid("theme replacement root range is invalid"))?;
                return optional_attribute_value_range(raw, wanted)
                    .map(|range| range.map(|range| (start + range.start)..(start + range.end)));
            },
            Event::DocType(_) | Event::PI(_) | Event::End(_) => {
                return Err(invalid(
                    "theme replacement has forbidden markup before its root",
                ));
            },
            Event::Eof => return Ok(None),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn optional_attribute_value_range(raw: &[u8], wanted: &[u8]) -> Result<Option<Range<usize>>> {
    let mut cursor = 1usize;
    while cursor < raw.len()
        && !raw[cursor].is_ascii_whitespace()
        && !matches!(raw[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    let mut found = None;
    while cursor < raw.len() {
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || matches!(raw[cursor], b'>' | b'/') {
            break;
        }
        let key_start = cursor;
        while cursor < raw.len()
            && !raw[cursor].is_ascii_whitespace()
            && !matches!(raw[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let key_end = cursor;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || raw[cursor] != b'=' {
            return Err(invalid("theme replacement attribute has no value"));
        }
        cursor += 1;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *raw
            .get(cursor)
            .ok_or_else(|| invalid("theme replacement attribute value is missing"))?;
        if !matches!(quote, b'"' | b'\'') {
            return Err(invalid("theme replacement attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < raw.len() && raw[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= raw.len() {
            return Err(invalid("theme replacement attribute value is unterminated"));
        }
        if &raw[key_start..key_end] == wanted {
            if found.is_some() {
                return Err(invalid(
                    "theme replacement has duplicate namespace attributes",
                ));
            }
            found = Some(value_start..value_end);
        }
        cursor += 1;
    }
    Ok(found)
}

fn direct_scheme_range(xml: &[u8], target: &[u8]) -> Result<Option<Range<usize>>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    if !matches!(target, b"clrScheme" | b"fontScheme") {
        return Err(invalid("unsupported theme scheme for replacement"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    let mut stack: Vec<(usize, Frame)> = Vec::new();
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut found = None;
    let mut target_namespace = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("theme XML offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("theme XML offset exceeds usize"))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        reject_unknown_namespace(&namespace)?;
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| invalid("theme node count overflow"))?;
        if nodes > MAX_NODES {
            return Err(limit("theme XML nodes", MAX_NODES));
        }
        match event {
            Event::Start(element) => {
                let local = element.local_name().as_ref().to_vec();
                let current_namespace = drawingml_namespace(&namespace);
                let drawingml = current_namespace.is_some();
                if !root_seen {
                    if local != b"theme" || !drawingml {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                } else if stack.is_empty() {
                    return Err(invalid("theme XML has multiple root elements"));
                }
                let direct =
                    is_direct_theme_parent(&stack) && drawingml && local.as_slice() == target;
                if direct {
                    if found.is_some() {
                        return Err(invalid("theme contains multiple replacement schemes"));
                    }
                    let source_namespace = current_namespace
                        .ok_or_else(|| invalid("theme scheme contains an unsupported namespace"))?;
                    validate_scheme_element(
                        target,
                        &stack,
                        &local,
                        current_namespace,
                        source_namespace,
                        &element,
                        reader.decoder(),
                    )?;
                    target_namespace = Some(source_namespace);
                } else if is_target_stack(&stack, target) {
                    let source_namespace = target_namespace
                        .ok_or_else(|| invalid("theme scheme namespace state is missing"))?;
                    validate_scheme_element(
                        target,
                        &stack,
                        &local,
                        current_namespace,
                        source_namespace,
                        &element,
                        reader.decoder(),
                    )?;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("theme XML depth", MAX_DEPTH));
                }
                stack.push((start, Frame { local, drawingml }));
            },
            Event::Empty(element) => {
                let local = element.local_name().as_ref().to_vec();
                let current_namespace = drawingml_namespace(&namespace);
                let drawingml = current_namespace.is_some();
                if !root_seen {
                    if local != b"theme" || !drawingml {
                        return Err(invalid("theme XML has an unexpected root"));
                    }
                    root_seen = true;
                } else if stack.is_empty() {
                    return Err(invalid("theme XML has multiple root elements"));
                }
                let direct =
                    is_direct_theme_parent(&stack) && drawingml && local.as_slice() == target;
                if direct || is_target_stack(&stack, target) {
                    let source_namespace = if direct {
                        current_namespace.ok_or_else(|| {
                            invalid("theme scheme contains an unsupported namespace")
                        })?
                    } else {
                        target_namespace
                            .ok_or_else(|| invalid("theme scheme namespace state is missing"))?
                    };
                    validate_scheme_element(
                        target,
                        &stack,
                        &local,
                        current_namespace,
                        source_namespace,
                        &element,
                        reader.decoder(),
                    )?;
                }
                if direct {
                    if found.is_some() {
                        return Err(invalid("theme contains multiple replacement schemes"));
                    }
                    found = Some(start..end);
                }
            },
            Event::End(element) => {
                let current_namespace = drawingml_namespace(&namespace);
                let drawingml = current_namespace.is_some();
                if is_target_stack(&stack, target) {
                    let source_namespace = target_namespace
                        .ok_or_else(|| invalid("theme scheme namespace state is missing"))?;
                    if current_namespace != Some(source_namespace) {
                        return Err(invalid("theme scheme contains a cross-family namespace"));
                    }
                }
                let Some((open, frame)) = stack.pop() else {
                    return Err(invalid("theme XML nesting underflow"));
                };
                if frame.local.as_slice() != element.local_name().as_ref()
                    || frame.drawingml != drawingml
                {
                    return Err(invalid("theme XML closing element does not match"));
                }
                if stack.len() == 2
                    && is_direct_theme_parent(&stack)
                    && frame.drawingml
                    && frame.local.as_slice() == target
                {
                    if found.is_some() {
                        return Err(invalid("theme contains multiple replacement schemes"));
                    }
                    found = Some(open..end);
                    target_namespace = None;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("theme XML contains forbidden markup"));
            },
            Event::Eof => break,
            Event::Text(text) if is_target_stack(&stack, target) => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("theme scheme contains unsupported text"));
                }
            },
            Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_)
                if is_target_stack(&stack, target) =>
            {
                return Err(invalid("theme scheme contains unsupported content"));
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !root_seen || !stack.is_empty() {
        return Err(invalid("theme XML is unterminated"));
    }
    Ok(found)
}

fn is_direct_theme_parent(stack: &[(usize, Frame)]) -> bool {
    stack.len() == 2
        && stack[0].1.drawingml
        && stack[0].1.local.as_slice() == b"theme"
        && stack[1].1.drawingml
        && stack[1].1.local.as_slice() == b"themeElements"
}

fn is_target_stack(stack: &[(usize, Frame)], target: &[u8]) -> bool {
    stack.len() >= 3
        && is_direct_theme_parent(&stack[..2])
        && stack[2].1.drawingml
        && stack[2].1.local.as_slice() == target
}

fn validate_scheme_element(
    target: &[u8],
    stack: &[(usize, Frame)],
    local: &[u8],
    current_namespace: Option<&[u8]>,
    source_namespace: &[u8],
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let current_namespace = current_namespace
        .ok_or_else(|| invalid("theme scheme contains an unsupported namespace"))?;
    if current_namespace != source_namespace {
        return Err(invalid("theme scheme contains an unsupported namespace"));
    }
    if stack.len() == 2 {
        if local != target {
            return Err(invalid("theme scheme contains an unsupported child"));
        }
        ensure_supported_attributes(element, decoder, &[b"name"], source_namespace)?;
        return Ok(());
    }
    let parent = stack
        .last()
        .map(|(_, frame)| frame.local.as_slice())
        .unwrap_or_default();
    if target == b"clrScheme" {
        if parent == b"clrScheme" {
            if Slot::from_token(std::str::from_utf8(local).unwrap_or_default()).is_none() {
                return Err(invalid("color scheme contains an unsupported child"));
            }
            ensure_supported_attributes(element, decoder, &[], source_namespace)?;
        } else if is_color_value_parent(stack) {
            match local {
                b"srgbClr" => {
                    ensure_supported_attributes(element, decoder, &[b"val"], source_namespace)?
                },
                b"sysClr" => {
                    ensure_supported_attributes(
                        element,
                        decoder,
                        &[b"val", b"lastClr"],
                        source_namespace,
                    )?;
                },
                _ => return Err(invalid("color slot contains an unsupported child")),
            }
        } else {
            return Err(invalid("color scheme contains an unsupported child"));
        }
    } else if target == b"fontScheme" {
        if parent == b"fontScheme" {
            if !matches!(local, b"majorFont" | b"minorFont") {
                return Err(invalid("font scheme contains an unsupported child"));
            }
            ensure_supported_attributes(element, decoder, &[], source_namespace)?;
        } else if is_font_face_parent_pair(stack) {
            let allowed = match local {
                b"latin" | b"ea" | b"cs" => &[&b"typeface"[..]][..],
                b"font" => &[&b"script"[..], &b"typeface"[..]][..],
                _ => return Err(invalid("font face contains an unsupported child")),
            };
            ensure_supported_attributes(element, decoder, allowed, source_namespace)?;
        } else {
            return Err(invalid("font scheme contains an unsupported child"));
        }
    }
    Ok(())
}

fn is_color_value_parent(stack: &[(usize, Frame)]) -> bool {
    stack.last().is_some_and(|(_, frame)| {
        frame.drawingml
            && Slot::from_token(std::str::from_utf8(&frame.local).unwrap_or_default()).is_some()
    })
}

fn is_font_face_parent_pair(stack: &[(usize, Frame)]) -> bool {
    stack.last().is_some_and(|(_, frame)| {
        frame.drawingml && matches!(frame.local.as_slice(), b"majorFont" | b"minorFont")
    })
}

fn ensure_supported_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    allowed: &[&[u8]],
    source_namespace: &[u8],
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            if value.as_ref().as_bytes() != source_namespace {
                return Err(invalid(
                    "theme scheme declares a foreign or cross-family namespace",
                ));
            }
            continue;
        }
        if attribute.key.prefix().is_some() || !allowed.contains(&key) {
            return Err(invalid("theme scheme contains an unsupported attribute"));
        }
        attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    Ok(())
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
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
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
    use crate::theme::model::{Color, Face, FontSet, Palette, Slot, System};

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
    fn all_system_color_tokens_round_trip() {
        assert_eq!(System::ALL.len(), 30);
        for kind in System::ALL {
            assert_eq!(System::from_token(kind.token()), Some(kind));
            assert!(Color::system(kind, Some("000000")).is_ok());
        }
        let colors = Slot::ALL
            .into_iter()
            .fold(Palette::new("Office"), |palette, slot| {
                palette.with(
                    slot,
                    Color::system(System::ThreeDDarkShadow, Some("000000")).unwrap(),
                )
            });
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let parsed = read(&encode_part("Office", &colors, &fonts).unwrap()).unwrap();
        assert_eq!(parsed.colors, colors);

        let invalid = String::from_utf8(encode_part("Office", &palette(), &fonts).unwrap())
            .unwrap()
            .replacen(
                "<a:srgbClr val=\"4F81BD\"/>",
                "<a:sysClr val=\"futureColor\" lastClr=\"000000\"/>",
                1,
            );
        assert!(read(invalid.as_bytes()).is_err());
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
        let color_range = scheme_replacement_range(&original, b"clrScheme").unwrap();
        assert!(color_range.start < color_range.end);
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

    #[test]
    fn authored_format_scheme_has_schema_required_entries() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let xml = encode_part("Office", &palette(), &fonts).unwrap();
        let text = String::from_utf8(xml).unwrap();
        assert_eq!(text.matches("<a:solidFill>").count(), 9);
        assert_eq!(text.matches("<a:ln ").count(), 3);
        assert_eq!(text.matches("<a:effectStyle>").count(), 3);
        assert_eq!(text.matches("<a:effectLst/>").count(), 3);
    }

    #[test]
    fn typed_read_requires_drawingml_namespaces_and_direct_scheme_parents() {
        let xml = br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:foreign" name="Foreign"><a:themeElements><x:clrScheme name="Foreign"><x:dk1><x:srgbClr val="000000"/></x:dk1></x:clrScheme><a:fontScheme name="Foreign"><a:majorFont><a:latin typeface="Aptos"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#;
        assert!(read(xml).is_err());

        let nested = br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:wrapper><a:clrScheme name="Office"/></a:wrapper></a:themeElements></a:theme>"#;
        assert!(replace_scheme(nested, b"clrScheme", b"<a:clrScheme/>").is_err());

        let undeclared = br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><x:clrScheme/></a:themeElements></a:theme>"#;
        assert!(read(undeclared).is_err());
    }

    #[test]
    fn scheme_replacement_refuses_ignored_source_content() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let original =
            String::from_utf8(encode_part("Office", &palette(), &fonts).unwrap()).unwrap();
        let with_unknown_child =
            original.replacen("</a:clrScheme>", "<a:extLst/></a:clrScheme>", 1);
        assert!(
            replace_scheme(
                with_unknown_child.as_bytes(),
                b"clrScheme",
                b"<a:clrScheme/>"
            )
            .is_err()
        );

        let with_unknown_attribute = original.replacen(
            "<a:clrScheme name=\"Office\">",
            "<a:clrScheme name=\"Office\" future=\"1\">",
            1,
        );
        assert!(
            replace_scheme(
                with_unknown_attribute.as_bytes(),
                b"clrScheme",
                b"<a:clrScheme/>"
            )
            .is_err()
        );

        let with_foreign_namespace = original.replacen(
            "<a:clrScheme name=\"Office\">",
            "<a:clrScheme xmlns:x=\"urn:foreign\" name=\"Office\">",
            1,
        );
        assert!(
            replace_scheme(
                with_foreign_namespace.as_bytes(),
                b"clrScheme",
                b"<a:clrScheme/>"
            )
            .is_err()
        );

        let with_cross_family_namespace = original.replacen(
            "<a:clrScheme name=\"Office\">",
            "<a:clrScheme xmlns:x=\"http://purl.oclc.org/ooxml/drawingml/main\" name=\"Office\">",
            1,
        );
        assert!(
            replace_scheme(
                with_cross_family_namespace.as_bytes(),
                b"clrScheme",
                b"<a:clrScheme/>"
            )
            .is_err()
        );

        let with_matching_namespace = original.replacen(
            "<a:clrScheme name=\"Office\">",
            "<a:clrScheme xmlns:x=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\">",
            1,
        );
        assert!(
            replace_scheme(
                with_matching_namespace.as_bytes(),
                b"clrScheme",
                b"<a:clrScheme/>"
            )
            .is_ok()
        );
    }

    #[test]
    fn scheme_replacement_rebinds_generated_fragment_to_strict_source() {
        let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        let original = encode_part("Office", &palette(), &fonts).unwrap();
        let strict = original
            .windows(NAMESPACE.len())
            .enumerate()
            .find_map(|(index, window)| (window == NAMESPACE.as_bytes()).then_some(index))
            .map(|index| {
                let mut bytes = original.clone();
                bytes.splice(
                    index..index + NAMESPACE.len(),
                    STRICT_NAMESPACE.as_bytes().iter().copied(),
                );
                bytes
            })
            .unwrap();
        let strict_name = Slot::ALL
            .into_iter()
            .fold(Palette::new(NAMESPACE), |palette, slot| {
                palette.with(slot, Color::rgb("FFFFFF").unwrap())
            });
        let patched = replace_scheme(
            &strict,
            b"clrScheme",
            &encode_palette_fragment(&strict_name).unwrap(),
        )
        .unwrap();
        assert!(
            patched
                .windows(STRICT_NAMESPACE.len())
                .any(|window| { window == STRICT_NAMESPACE.as_bytes() })
        );
        assert!(
            patched
                .windows(NAMESPACE.len())
                .any(|window| { window == NAMESPACE.as_bytes() })
        );
        assert!(
            patched
                .windows(b"xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\"".len())
                .any(|window| {
                    window == b"xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\""
                })
        );
        assert_eq!(read(&patched).unwrap().colors.name(), NAMESPACE);
    }

    #[test]
    fn authored_text_rejects_xml_forbidden_chars_and_preserves_whitespace() {
        let colors = Slot::ALL
            .into_iter()
            .fold(Palette::new("Office\n"), |palette, slot| {
                palette.with(slot, Color::rgb("4472C4").unwrap())
            });
        let fonts = FontSet::new(
            "Office\t",
            Face::new("Aptos\r")
                .east_asian("Noto\n")
                .complex_script("Arabic\t"),
            Face::new("Aptos").script("Latn", "Noto Sans"),
        );
        let xml = encode_part("Theme\n", &colors, &fonts).unwrap();
        assert!(
            xml.windows(b"name=\"Theme&#xA;\"".len())
                .any(|window| window == b"name=\"Theme&#xA;\"")
        );
        assert!(
            xml.windows(b"name=\"Office&#x9;\"".len())
                .any(|window| window == b"name=\"Office&#x9;\"")
        );
        assert!(
            xml.windows(b"typeface=\"Aptos&#xD;\"".len())
                .any(|window| window == b"typeface=\"Aptos&#xD;\"")
        );
        let parsed = read(&xml).unwrap();
        assert_eq!(parsed.name, "Theme\n");
        assert_eq!(parsed.fonts.name(), "Office\t");
        assert_eq!(parsed.fonts.major().latin, "Aptos\r");
        assert_eq!(parsed.fonts.major().east_asian, "Noto\n");
        assert_eq!(parsed.fonts.major().complex_script, "Arabic\t");

        let bad_palette = Slot::ALL
            .into_iter()
            .fold(Palette::new("bad\u{1}"), |palette, slot| {
                palette.with(slot, Color::rgb("4472C4").unwrap())
            });
        let plain_fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
        assert!(encode_part("Office", &bad_palette, &plain_fonts).is_err());
        assert!(
            encode_part(
                "Office",
                &palette(),
                &FontSet::new("Office", Face::new("Aptos\u{B}"), Face::new("Aptos"))
            )
            .is_err()
        );
        assert!(
            encode_part(
                "Office",
                &palette(),
                &FontSet::new(
                    "Office",
                    Face::new("Aptos").script("Latn", "Noto\u{1}"),
                    Face::new("Aptos")
                )
            )
            .is_err()
        );
    }
}
