//! Bounded PresentationML and font-container codecs.

use super::model::{Charset, Conformance, License, Panose, PitchFamily, Style};
use super::{
    FONT_DATA_CT, FONT_REL, FONT_TTF_CT, MAX_DEPTH, MAX_FONT_BYTES, MAX_FONTS, MAX_NODES,
    MAX_STRING_BYTES, MAX_TOTAL_FONT_BYTES, MAX_XML_BYTES, PML, STRICT_FONT_REL, STRICT_PML,
    invalid, limit,
};
use crate::error::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::PackURI;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashMap;
use std::sync::Arc;

#[derive(Debug, Clone)]
pub(super) struct RawResource {
    pub(super) part_name: String,
    pub(super) content_type: String,
    /// The font program is deliberately retained as inert bytes.
    pub(super) data: Arc<Vec<u8>>,
}

impl PartialEq for RawResource {
    fn eq(&self, other: &Self) -> bool {
        self.part_name == other.part_name
            && self.content_type == other.content_type
            && (Arc::ptr_eq(&self.data, &other.data) || self.data == other.data)
    }
}

impl Eq for RawResource {}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct RawFace {
    pub(super) style: Style,
    pub(super) relationship_id: String,
    /// Present after package loading and required for package storage.
    pub(super) resource: Option<RawResource>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct RawFont {
    pub(super) has_descriptor: bool,
    pub(super) typeface: String,
    pub(super) panose: Option<Panose>,
    pub(super) pitch_family: Option<PitchFamily>,
    pub(super) charset: Option<Charset>,
    pub(super) faces: Vec<RawFace>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(super) struct RawFonts {
    pub(super) fonts: Vec<RawFont>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Presentation,
    List,
    Font(usize),
    Leaf,
    Other,
}

pub(super) struct ParsedPresentation {
    pub(super) conformance: Conformance,
    pub(super) value: Option<RawFonts>,
}

/// Parse the optional embedded-font markup from a complete presentation part.
#[cfg(test)]
pub(super) fn parse_raw(xml: &[u8]) -> Result<Option<RawFonts>> {
    Ok(parse_presentation(xml)?.value)
}

pub(super) fn parse_presentation(xml: &[u8]) -> Result<ParsedPresentation> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("presentation XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_XML_BYTES {
        return Err(limit("MCE-processed presentation XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Context> = Vec::new();
    let mut fonts = Vec::new();
    let mut saw_root = false;
    let mut saw_list = false;
    let mut conformance = None;
    let mut root_rank = None;
    let mut nodes = 0usize;
    let mut string_bytes = 0usize;
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        let empty_event = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = empty_event;
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
                let namespace =
                    resolved_namespace(reader.resolver().resolve_element(element.name()).0)?;
                let local = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                let parent = stack.last().copied();
                if parent == Some(Context::Presentation)
                    && namespace == conformance.map(Conformance::pml).unwrap_or_default()
                    && let Some(rank) = presentation_child_rank(&local)
                {
                    if root_rank.is_some_and(|previous| previous > rank) {
                        return Err(invalid(format!(
                            "presentation child '{local}' is out of schema order"
                        )));
                    }
                    root_rank = Some(rank);
                }
                let context = if parent.is_none() {
                    if saw_root || local != "presentation" {
                        return Err(invalid("expected one PresentationML presentation root"));
                    }
                    let c = match namespace.as_str() {
                        PML => Conformance::Transitional,
                        STRICT_PML => Conformance::Strict,
                        _ => return Err(invalid("presentation root has an unsupported namespace")),
                    };
                    saw_root = true;
                    conformance = Some(c);
                    Context::Presentation
                } else if parent == Some(Context::Presentation)
                    && namespace == conformance.map(Conformance::pml).unwrap_or_default()
                    && local == "embeddedFontLst"
                {
                    if saw_list {
                        return Err(invalid(
                            "presentation has multiple embeddedFontLst elements",
                        ));
                    }
                    reject_unqualified_attributes(
                        &reader,
                        element,
                        reader.decoder(),
                        &[],
                        &mut string_bytes,
                    )?;
                    saw_list = true;
                    Context::List
                } else if parent == Some(Context::List) {
                    if namespace != conformance.map(Conformance::pml).unwrap_or_default()
                        || local != "embeddedFont"
                    {
                        return Err(invalid("embeddedFontLst contains a non-embeddedFont child"));
                    }
                    reject_unqualified_attributes(
                        &reader,
                        element,
                        reader.decoder(),
                        &[],
                        &mut string_bytes,
                    )?;
                    if fonts.len() >= MAX_FONTS {
                        return Err(limit("embedded fonts"));
                    }
                    fonts.push(RawFont {
                        has_descriptor: false,
                        typeface: String::new(),
                        panose: None,
                        pitch_family: None,
                        charset: None,
                        faces: Vec::new(),
                    });
                    Context::Font(fonts.len() - 1)
                } else if let Some(Context::Font(index)) = parent {
                    if namespace != conformance.map(Conformance::pml).unwrap_or_default() {
                        return Err(invalid("embeddedFont contains a foreign child"));
                    }
                    if local == "font" {
                        if fonts[index].has_descriptor {
                            return Err(invalid("embeddedFont has multiple font descriptors"));
                        }
                        if !fonts[index].faces.is_empty() {
                            return Err(invalid(
                                "embeddedFont descriptor must precede every style face",
                            ));
                        }
                        parse_descriptor(
                            &reader,
                            element,
                            reader.decoder(),
                            &mut fonts[index],
                            &mut string_bytes,
                        )?;
                        fonts[index].has_descriptor = true;
                        Context::Leaf
                    } else if let Some(style) = Style::parse_raw(&local) {
                        if !fonts[index].has_descriptor {
                            return Err(invalid(
                                "embeddedFont descriptor must precede every style face",
                            ));
                        }
                        let relationship_id = parse_face(
                            &reader,
                            element,
                            reader.decoder(),
                            conformance.ok_or_else(|| invalid("missing presentation profile"))?,
                            &mut string_bytes,
                        )?;
                        if fonts[index].faces.iter().any(|face| face.style == style) {
                            return Err(invalid(format!(
                                "duplicate embedded-font style '{local}'"
                            )));
                        }
                        if fonts[index]
                            .faces
                            .last()
                            .is_some_and(|face| face.style.rank() >= style.rank())
                        {
                            return Err(invalid("embedded-font styles are out of schema order"));
                        }
                        fonts[index].faces.push(RawFace {
                            style,
                            relationship_id,
                            resource: None,
                        });
                        Context::Leaf
                    } else {
                        return Err(invalid(format!("unexpected embeddedFont child '{local}'")));
                    }
                } else if matches!(parent, Some(Context::List | Context::Leaf)) {
                    return Err(invalid("embedded-font leaf element contains child content"));
                } else {
                    Context::Other
                };
                stack.push(context);
                if empty {
                    let ended = stack
                        .pop()
                        .ok_or_else(|| invalid("missing empty-element context"))?;
                    if let Context::Font(index) = ended {
                        finish_font(&fonts[index])?;
                    }
                }
            },
            Event::End(_) => {
                let ended = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                if let Context::Font(index) = ended {
                    finish_font(&fonts[index])?;
                }
            },
            Event::Text(text) => {
                if matches!(
                    stack.last(),
                    Some(Context::List | Context::Font(_) | Context::Leaf)
                ) {
                    let value = text.decode().map_err(xml_error)?;
                    if !value.trim().is_empty() {
                        return Err(invalid("embedded-font markup contains text"));
                    }
                } else if stack.is_empty() && !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("text occurs outside the presentation root"));
                }
            },
            Event::CData(_) => {
                if matches!(
                    stack.last(),
                    Some(Context::List | Context::Font(_) | Context::Leaf)
                ) {
                    return Err(invalid("CDATA is rejected in embedded-font markup"));
                }
            },
            Event::GeneralRef(_) => {
                if matches!(
                    stack.last(),
                    Some(Context::List | Context::Font(_) | Context::Leaf)
                ) {
                    return Err(invalid(
                        "entity references are rejected in embedded-font markup",
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation XML"));
    }
    let conformance = conformance.ok_or_else(|| invalid("missing presentation root"))?;
    let value = saw_list.then_some(RawFonts { fonts });
    if let Some(value) = &value {
        validate_value(value, false)?;
    }
    Ok(ParsedPresentation { conformance, value })
}

pub(super) fn parse_descriptor(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    font: &mut RawFont,
    strings: &mut usize,
) -> Result<()> {
    let attrs = collect_unqualified_attributes(
        reader,
        element,
        decoder,
        &["typeface", "panose", "pitchFamily", "charset"],
        strings,
    )?;
    font.typeface = attrs
        .get("typeface")
        .cloned()
        .ok_or_else(|| invalid("font descriptor is missing typeface"))?;
    font.panose = attrs
        .get("panose")
        .map(|value| parse_panose(value))
        .transpose()?;
    font.pitch_family = attrs
        .get("pitchFamily")
        .map(|value| {
            value
                .parse::<u8>()
                .map_err(|_| invalid(format!("invalid pitchFamily value '{value}'")))
                .and_then(PitchFamily::from_wire)
        })
        .transpose()?;
    font.charset = attrs
        .get("charset")
        .map(|value| {
            value
                .parse::<i8>()
                .map_err(|_| invalid(format!("invalid charset byte value '{value}'")))
                .map(Charset::from_wire)
        })
        .transpose()?;
    Ok(())
}

pub(super) fn parse_face(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    conformance: Conformance,
    strings: &mut usize,
) -> Result<String> {
    let mut id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let qualified = attribute.key.as_ref();
        if qualified == b"xmlns" || qualified.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() != b"id"
            || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == conformance.rel_ns().as_bytes())
        {
            return Err(invalid("embedded-font face has an unexpected attribute"));
        }
        if id.is_some() {
            return Err(invalid("embedded-font face has duplicate relationship IDs"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_string_bytes(strings, value.len())?;
        id = Some(value);
    }
    let id = id
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid("embedded-font face is missing r:id"))?;
    validate_relationship_id(&id)?;
    Ok(id)
}

pub(super) fn collect_unqualified_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    allowed: &[&str],
    strings: &mut usize,
) -> Result<HashMap<String, String>> {
    let mut result = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let qualified = attribute.key.as_ref();
        if qualified == b"xmlns" || qualified.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let local = std::str::from_utf8(local.as_ref()).map_err(xml_error)?;
        if namespace != ResolveResult::Unbound || !allowed.contains(&local) {
            return Err(invalid(format!("unexpected attribute '{local}'")));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_string_bytes(strings, local.len() + value.len())?;
        if result.insert(local.to_owned(), value).is_some() {
            return Err(invalid(format!("duplicate attribute '{local}'")));
        }
    }
    Ok(result)
}

pub(super) fn reject_unqualified_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    allowed: &[&str],
    strings: &mut usize,
) -> Result<()> {
    collect_unqualified_attributes(reader, element, decoder, allowed, strings).map(|_| ())
}

/// Deterministically serializes a self-contained `p:embeddedFontLst` fragment.
pub(super) fn write_raw(value: &RawFonts, conformance: Conformance) -> Result<Vec<u8>> {
    validate_value(value, false)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<p:embeddedFontLst xmlns:p=\"");
    escape(&mut output, conformance.pml());
    output.extend_from_slice(b"\" xmlns:r=\"");
    escape(&mut output, conformance.rel_ns());
    if value.fonts.is_empty() {
        output.extend_from_slice(b"\"/>");
        return Ok(output);
    }
    output.extend_from_slice(b"\">");
    for font in &value.fonts {
        output.extend_from_slice(b"<p:embeddedFont><p:font typeface=\"");
        escape(&mut output, &font.typeface);
        output.push(b'\"');
        if let Some(panose) = font.panose {
            attribute(&mut output, "panose", &hex_panose(panose)?);
        }
        if let Some(value) = font.pitch_family {
            attribute(&mut output, "pitchFamily", &value.wire().to_string());
        }
        if let Some(value) = font.charset {
            attribute(&mut output, "charset", &value.wire().to_string());
        }
        output.extend_from_slice(b"/>");
        for face in &font.faces {
            output.extend_from_slice(b"<p:");
            output.extend_from_slice(face.style.element().as_bytes());
            output.extend_from_slice(b" r:id=\"");
            escape(&mut output, &face.relationship_id);
            output.extend_from_slice(b"\"/>");
        }
        output.extend_from_slice(b"</p:embeddedFont>");
    }
    output.extend_from_slice(b"</p:embeddedFontLst>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized embedded-font XML bytes"));
    }
    Ok(output)
}

pub(super) fn presentation_child_rank(local: &str) -> Option<usize> {
    [
        "sldMasterIdLst",
        "notesMasterIdLst",
        "handoutMasterIdLst",
        "sldIdLst",
        "sldSz",
        "notesSz",
        "smartTags",
        "embeddedFontLst",
        "custShowLst",
        "photoAlbum",
        "custDataLst",
        "kinsoku",
        "defaultTextStyle",
        "modifyVerifier",
        "extLst",
    ]
    .iter()
    .position(|name| *name == local)
}

pub(super) fn validate_typeface(value: &str) -> Result<()> {
    bounded_string(value)?;
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid(
            "embedded-font typeface contains an XML 1.0-forbidden character",
        ))
    }
}

pub(super) fn validate_font_bytes(value: &[u8]) -> Result<()> {
    if value.len() > MAX_FONT_BYTES {
        Err(limit("individual font bytes"))
    } else {
        Ok(())
    }
}

pub(super) fn validate_eot(value: &[u8]) -> Result<()> {
    const VERSION_1: u32 = 0x0001_0000;
    const VERSION_2_1: u32 = 0x0002_0001;
    const VERSION_2_2: u32 = 0x0002_0002;
    const FLAGS: u32 = 0x1000_00F5;
    const EUDC: u32 = 0x0000_0020;

    let eot_size = usize::try_from(le_u32(value, 0)?)
        .map_err(|_| invalid("EOT size does not fit this platform"))?;
    if eot_size != value.len() {
        return Err(invalid("EOT size does not match the container length"));
    }
    let font_size = usize::try_from(le_u32(value, 4)?)
        .map_err(|_| invalid("EOT font-data size does not fit this platform"))?;
    if font_size == 0 {
        return Err(invalid("EOT font-data payload is empty"));
    }
    let font_start = value
        .len()
        .checked_sub(font_size)
        .ok_or_else(|| invalid("EOT font-data size exceeds the container"))?;
    let version = le_u32(value, 8)?;
    if !matches!(version, VERSION_1 | VERSION_2_1 | VERSION_2_2) {
        return Err(invalid(format!("unsupported EOT version 0x{version:08X}")));
    }
    let flags = le_u32(value, 12)?;
    if flags & !FLAGS != 0 {
        return Err(invalid(format!(
            "EOT processing flags contain reserved bits 0x{:08X}",
            flags & !FLAGS
        )));
    }
    if version == VERSION_1 && flags & EUDC != 0 {
        return Err(invalid("EOT version 1 cannot contain an EUDC font"));
    }
    if value.get(27).copied().is_none_or(|italic| italic > 1) {
        return Err(invalid("EOT italic byte must be zero or one"));
    }
    License::from_fs_type(le_u16(value, 32)?)?;
    if le_u16(value, 34)? != 0x504C {
        return Err(invalid("EOT magic number is not 0x504C"));
    }
    if value
        .get(64..80)
        .is_none_or(|reserved| reserved.iter().any(|byte| *byte != 0))
    {
        return Err(invalid("EOT reserved header words must be zero"));
    }
    if le_u16(value, 80)? != 0 {
        return Err(invalid("EOT header padding must be zero"));
    }

    let mut cursor = 82usize;
    for name in ["family", "style", "version", "full"] {
        eot_utf16(value, &mut cursor, font_start, name)?;
        if name != "full" && eot_u16(value, &mut cursor, font_start, "name padding")? != 0 {
            return Err(invalid("EOT name padding must be zero"));
        }
    }

    if version != VERSION_1 {
        if eot_u16(value, &mut cursor, font_start, "root padding")? != 0 {
            return Err(invalid("EOT root-string padding must be zero"));
        }
        let root = eot_sized(value, &mut cursor, font_start, "root string")?;
        if root.len() % 2 != 0 {
            return Err(invalid("EOT root string is not UTF-16 byte-aligned"));
        }
        if version == VERSION_2_2 {
            let checksum = eot_u32(value, &mut cursor, font_start, "root checksum")?;
            let expected = root
                .iter()
                .fold(0u32, |sum, byte| sum.wrapping_add(u32::from(*byte)))
                ^ 0x5047_5342;
            if checksum != expected {
                return Err(invalid("EOT root-string checksum is invalid"));
            }
            let _code_page = eot_u32(value, &mut cursor, font_start, "EUDC code page")?;
            if eot_u16(value, &mut cursor, font_start, "signature padding")? != 0 {
                return Err(invalid("EOT signature padding must be zero"));
            }
            let signature = eot_sized(value, &mut cursor, font_start, "signature")?;
            if !signature.is_empty() {
                return Err(invalid("EOT reserved signature must be empty"));
            }
            let eudc_flags = eot_u32(value, &mut cursor, font_start, "EUDC flags")?;
            if eudc_flags & !FLAGS != 0 {
                return Err(invalid("EOT EUDC flags contain reserved bits"));
            }
            let eudc_size = usize::try_from(eot_u32(
                value,
                &mut cursor,
                font_start,
                "EUDC font-data size",
            )?)
            .map_err(|_| invalid("EOT EUDC font-data size does not fit this platform"))?;
            eot_take(value, &mut cursor, eudc_size, font_start, "EUDC font data")?;
            if (flags & EUDC != 0) != (eudc_size != 0) {
                return Err(invalid("EOT EUDC flag and payload disagree"));
            }
        }
    }
    if cursor != font_start {
        return Err(invalid(
            "EOT variable header overlaps or precedes font data",
        ));
    }
    if flags & (0x0000_0004 | 0x1000_0000) == 0 {
        validate_sfnt(
            value
                .get(font_start..)
                .ok_or_else(|| invalid("EOT font-data range is invalid"))?,
        )?;
    }
    Ok(())
}

pub(super) fn validate_sfnt(value: &[u8]) -> Result<()> {
    match value.get(..4) {
        Some(b"ttcf") => {
            let version = be_u32(value, 4)?;
            if !matches!(version, 0x0001_0000 | 0x0002_0000) {
                return Err(invalid(format!(
                    "unsupported TrueType Collection version 0x{version:08X}"
                )));
            }
            let fonts = usize::try_from(be_u32(value, 8)?)
                .map_err(|_| invalid("TrueType Collection font count does not fit"))?;
            if fonts == 0 {
                return Err(invalid("TrueType Collection contains no fonts"));
            }
            let offsets_end = 12usize
                .checked_add(
                    fonts
                        .checked_mul(4)
                        .ok_or_else(|| invalid("TrueType Collection offset table overflows"))?,
                )
                .ok_or_else(|| invalid("TrueType Collection offset table overflows"))?;
            if offsets_end > value.len() {
                return Err(invalid("TrueType Collection offset table is truncated"));
            }
            for index in 0..fonts {
                let field = 12usize
                    .checked_add(index * 4)
                    .ok_or_else(|| invalid("TrueType Collection font offset overflows"))?;
                let offset = usize::try_from(be_u32(value, field)?)
                    .map_err(|_| invalid("TrueType Collection font offset does not fit"))?;
                if offset % 4 != 0 {
                    return Err(invalid("TrueType Collection font offset is not aligned"));
                }
                validate_sfnt_at(value, offset)?;
            }
            Ok(())
        },
        Some(_) => validate_sfnt_at(value, 0),
        None => Err(invalid("sfnt container is missing its signature")),
    }
}

pub(super) fn validate_sfnt_at(value: &[u8], offset: usize) -> Result<()> {
    let signature_end = offset
        .checked_add(4)
        .ok_or_else(|| invalid("sfnt offset overflows"))?;
    if !matches!(
        value.get(offset..signature_end),
        Some(b"\0\x01\0\0" | b"OTTO" | b"true" | b"typ1")
    ) {
        return Err(invalid("font data has an unsupported sfnt signature"));
    }
    let tables = usize::from(be_u16(value, offset + 4)?);
    let directory_end = offset
        .checked_add(12)
        .and_then(|base| {
            tables
                .checked_mul(16)
                .and_then(|size| base.checked_add(size))
        })
        .ok_or_else(|| invalid("sfnt table directory overflows"))?;
    if directory_end > value.len() {
        return Err(invalid("sfnt table directory is truncated"));
    }
    Ok(())
}

pub(super) fn eot_utf16(value: &[u8], cursor: &mut usize, limit: usize, name: &str) -> Result<()> {
    let data = eot_sized(value, cursor, limit, name)?;
    if data.len() % 2 != 0 {
        return Err(invalid(format!(
            "EOT {name} name is not UTF-16 byte-aligned"
        )));
    }
    let words = data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]));
    if char::decode_utf16(words).any(|character| character.is_err()) {
        return Err(invalid(format!("EOT {name} name contains invalid UTF-16")));
    }
    Ok(())
}

pub(super) fn eot_sized<'a>(
    value: &'a [u8],
    cursor: &mut usize,
    limit: usize,
    name: &str,
) -> Result<&'a [u8]> {
    let size = usize::from(eot_u16(value, cursor, limit, name)?);
    eot_take(value, cursor, size, limit, name)
}

pub(super) fn eot_take<'a>(
    value: &'a [u8],
    cursor: &mut usize,
    size: usize,
    limit: usize,
    name: &str,
) -> Result<&'a [u8]> {
    let end = cursor
        .checked_add(size)
        .ok_or_else(|| invalid(format!("EOT {name} size overflows")))?;
    if end > limit {
        return Err(invalid(format!("EOT {name} extends into font data")));
    }
    let data = value
        .get(*cursor..end)
        .ok_or_else(|| invalid(format!("EOT {name} is truncated")))?;
    *cursor = end;
    Ok(data)
}

pub(super) fn eot_u16(value: &[u8], cursor: &mut usize, limit: usize, name: &str) -> Result<u16> {
    let bytes = eot_take(value, cursor, 2, limit, name)?;
    let bytes = <[u8; 2]>::try_from(bytes).map_err(xml_error)?;
    Ok(u16::from_le_bytes(bytes))
}

pub(super) fn eot_u32(value: &[u8], cursor: &mut usize, limit: usize, name: &str) -> Result<u32> {
    let bytes = eot_take(value, cursor, 4, limit, name)?;
    let bytes = <[u8; 4]>::try_from(bytes).map_err(xml_error)?;
    Ok(u32::from_le_bytes(bytes))
}

pub(super) fn le_u16(value: &[u8], offset: usize) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid("font header offset overflows"))?;
    let bytes = value
        .get(offset..end)
        .ok_or_else(|| invalid("font header is truncated"))?;
    Ok(u16::from_le_bytes(
        <[u8; 2]>::try_from(bytes).map_err(xml_error)?,
    ))
}

pub(super) fn le_u32(value: &[u8], offset: usize) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| invalid("font header offset overflows"))?;
    let bytes = value
        .get(offset..end)
        .ok_or_else(|| invalid("font header is truncated"))?;
    Ok(u32::from_le_bytes(
        <[u8; 4]>::try_from(bytes).map_err(xml_error)?,
    ))
}

pub(super) fn be_u16(value: &[u8], offset: usize) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid("font header offset overflows"))?;
    let bytes = value
        .get(offset..end)
        .ok_or_else(|| invalid("font header is truncated"))?;
    Ok(u16::from_be_bytes(
        <[u8; 2]>::try_from(bytes).map_err(xml_error)?,
    ))
}

pub(super) fn be_u32(value: &[u8], offset: usize) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| invalid("font header offset overflows"))?;
    let bytes = value
        .get(offset..end)
        .ok_or_else(|| invalid("font header is truncated"))?;
    Ok(u32::from_be_bytes(
        <[u8; 4]>::try_from(bytes).map_err(xml_error)?,
    ))
}

pub(super) fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&value)
        || ('\u{E000}'..='\u{FFFD}').contains(&value)
        || ('\u{10000}'..='\u{10FFFF}').contains(&value)
}

pub(super) fn validate_value(value: &RawFonts, require_resources: bool) -> Result<()> {
    if value.fonts.len() > MAX_FONTS {
        return Err(limit("embedded fonts"));
    }
    let mut resources = HashMap::<&str, (&str, &Arc<Vec<u8>>)>::new();
    let mut total = 0usize;
    for font in &value.fonts {
        bounded_string(&font.typeface)?;
        finish_font(font)?;
        for face in &font.faces {
            validate_relationship_id(&face.relationship_id)?;
            if require_resources && face.resource.is_none() {
                return Err(invalid(
                    "embedded-font resource is required for package storage",
                ));
            }
            if let Some(resource) = &face.resource {
                PackURI::new(&resource.part_name).map_err(Error::Invalid)?;
                if !is_font_content_type(&resource.content_type) {
                    return Err(invalid(format!(
                        "invalid embedded-font content type '{}'",
                        resource.content_type
                    )));
                }
                validate_font_bytes(&resource.data)?;
                if let Some((content_type, data)) = resources.get(resource.part_name.as_str()) {
                    if *content_type != resource.content_type
                        || (!Arc::ptr_eq(data, &resource.data)
                            && data.as_slice() != resource.data.as_slice())
                    {
                        return Err(invalid(format!(
                            "shared font part '{}' has conflicting resources",
                            resource.part_name
                        )));
                    }
                } else {
                    resources.insert(
                        resource.part_name.as_str(),
                        (resource.content_type.as_str(), &resource.data),
                    );
                    total = total
                        .checked_add(resource.data.len())
                        .ok_or_else(|| limit("total font bytes"))?;
                    if total > MAX_TOTAL_FONT_BYTES {
                        return Err(limit("total font bytes"));
                    }
                }
            }
        }
    }
    Ok(())
}

pub(super) fn finish_font(font: &RawFont) -> Result<()> {
    if !font.has_descriptor {
        return Err(invalid("embeddedFont is missing its font descriptor"));
    }
    if font.faces.len() > 4 {
        return Err(invalid("embeddedFont has more than four styles"));
    }
    for pair in font.faces.windows(2) {
        if pair[0].style.rank() >= pair[1].style.rank() {
            return Err(invalid(
                "embedded-font styles are duplicated or out of schema order",
            ));
        }
    }
    Ok(())
}

pub(super) fn parse_panose(value: &str) -> Result<Panose> {
    if value.len() != 20 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid("panose must contain exactly 20 hexadecimal digits"));
    }
    let mut output = [0u8; 10];
    for (index, byte) in output.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16).map_err(xml_error)?;
    }
    Ok(Panose::new(output))
}

pub(super) fn hex_panose(value: Panose) -> Result<String> {
    let mut output = String::with_capacity(20);
    for byte in value.bytes() {
        use std::fmt::Write;
        write!(&mut output, "{byte:02X}").map_err(|_| Error::Write)?;
    }
    Ok(output)
}

pub(super) fn is_font_relationship(value: &str) -> bool {
    matches!(value, FONT_REL | STRICT_FONT_REL)
}
pub(super) fn is_font_content_type(value: &str) -> bool {
    matches!(value, FONT_DATA_CT | FONT_TTF_CT)
}
pub(super) fn bounded_string(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("embedded-font string bytes"))
    }
}
pub(super) fn add_string_bytes(total: &mut usize, count: usize) -> Result<()> {
    *total = total
        .checked_add(count)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
pub(super) fn validate_relationship_id(value: &str) -> Result<()> {
    if !litchi_ooxml_common::xml::is_ncname(value) {
        return Err(invalid(format!("invalid relationship ID '{value}'")));
    }
    Ok(())
}
pub(super) fn resolved_namespace(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
pub(super) fn attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'\"');
}
pub(super) fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
