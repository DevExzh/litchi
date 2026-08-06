//! Streaming WordprocessingML drawing inventory decoder.

use super::model::{Anchor, AnchorId, Kind, LegacyAnchor, LegacyAnchorKind, Object};
use super::validation::{parse_anchor_id_text, parse_word2010_anchor_id};
use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use litchi_core::unit::EMUS_PER_INCH;
use litchi_drawingml::geom::Preset;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use smallvec::SmallVec;

/// Parse drawing objects from a paragraph XML slice.
pub(crate) fn parse(xml_bytes: &[u8]) -> Result<SmallVec<[Object; 4]>> {
    let mut reader = Reader::from_reader(xml_bytes);
    reader.config_mut().trim_text(true);

    let mut objects = SmallVec::new();
    let mut in_drawing = false;
    let mut in_shape = false;
    let mut has_shape = false;
    let mut in_text_box = false;
    let mut in_text_content = false;

    let mut width_emu = EMUS_PER_INCH;
    let mut height_emu = EMUS_PER_INCH;
    let mut x_emu = 0;
    let mut y_emu = 0;
    let mut description = String::new();
    let mut name = String::new();
    let mut preset = None;
    let mut has_text_box = false;
    let mut is_inline = true;
    let mut anchor_id = None;
    let mut text_content = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                let local = element.local_name();
                match local.as_ref() {
                    b"drawing" => {
                        in_drawing = true;
                        width_emu = EMUS_PER_INCH;
                        height_emu = EMUS_PER_INCH;
                        x_emu = 0;
                        y_emu = 0;
                        description.clear();
                        name.clear();
                        preset = None;
                        has_text_box = false;
                        has_shape = false;
                        is_inline = true;
                        anchor_id = None;
                        text_content.clear();
                    },
                    b"inline" if in_drawing => {
                        is_inline = true;
                        anchor_id = parse_anchor_id(&element)?;
                    },
                    b"anchor" if in_drawing => {
                        is_inline = false;
                        anchor_id = parse_anchor_id(&element)?;
                    },
                    b"extent" if in_drawing => {
                        width_emu = number_attribute(&element, b"cx", EMUS_PER_INCH)?;
                        height_emu = number_attribute(&element, b"cy", EMUS_PER_INCH)?;
                    },
                    b"off" if in_drawing => {
                        x_emu = number_attribute(&element, b"x", 0)?;
                        y_emu = number_attribute(&element, b"y", 0)?;
                    },
                    b"ext" if in_drawing => {
                        width_emu = number_attribute(&element, b"cx", EMUS_PER_INCH)?;
                        height_emu = number_attribute(&element, b"cy", EMUS_PER_INCH)?;
                    },
                    b"docPr" if in_drawing => {
                        if let Some(value) = inert_attribute(&element, b"name") {
                            name = value;
                        }
                        if let Some(value) = inert_attribute(&element, b"descr") {
                            description = value;
                        }
                    },
                    b"wsp" if in_drawing => {
                        in_shape = true;
                        has_shape = true;
                    },
                    b"prstGeom" if in_shape => {
                        preset = Some(parse_preset(&element)?);
                    },
                    b"txbx" if in_shape => {
                        in_text_box = true;
                        has_text_box = true;
                    },
                    b"txbxContent" if in_text_box => {
                        in_text_content = true;
                    },
                    b"t" if in_text_content => {},
                    _ => {},
                }
            },
            Ok(Event::Text(text)) if in_text_content => {
                // Keep the legacy inventory behavior: malformed text bytes
                // are inert and do not prevent the surrounding drawing from
                // being discovered. The full Word story parser owns strict
                // text decoding elsewhere.
                if let Ok(value) = std::str::from_utf8(text.as_ref()) {
                    text_content.push_str(value);
                }
            },
            Ok(Event::End(element)) => {
                let local = element.local_name();
                match local.as_ref() {
                    b"drawing" => {
                        in_drawing = false;
                        let kind = if has_text_box {
                            Kind::TextBox
                        } else if has_shape {
                            Kind::Shape
                        } else {
                            Kind::Other
                        };
                        objects.push(Object::from_inventory(
                            name.clone(),
                            description.clone(),
                            width_emu,
                            height_emu,
                            x_emu,
                            y_emu,
                            preset,
                            kind,
                            if is_inline {
                                Anchor::Inline
                            } else {
                                Anchor::Floating
                            },
                            anchor_id,
                            text_content.clone(),
                        ));
                    },
                    b"wsp" => {
                        in_shape = false;
                    },
                    b"txbx" => {
                        in_text_box = false;
                    },
                    b"txbxContent" => {
                        in_text_content = false;
                    },
                    _ => {},
                }
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }

    Ok(objects)
}

/// Parse legacy Word `w:object` and `w:pict` anchors from a paragraph XML
/// slice. Only the checked `anchorId` metadata is projected; child VML, OLE,
/// and image content remains inert.
pub(crate) fn parse_legacy(xml_bytes: &[u8]) -> Result<SmallVec<[LegacyAnchor; 4]>> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut fragment_prefix = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut active = None;
    let mut anchors = SmallVec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("legacy drawing element counter overflow".to_string())
            })?;
            if nodes > 1_000_000 {
                return Err(Error::InvalidFormat(
                    "Word XML exceeds 1000000 elements".to_string(),
                ));
            }
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("legacy drawing XML nesting is too deep".to_string())
                })?;
                if depth > 128 {
                    return Err(Error::InvalidFormat(
                        "legacy drawing XML nesting exceeds 128".to_string(),
                    ));
                }
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix = Some(
                        element
                            .name()
                            .prefix()
                            .map(|prefix| prefix.into_inner().to_vec()),
                    );
                }
                if active.is_none()
                    && is_word_element(&namespace, &fragment_prefix)
                    && let Some(kind) = legacy_kind(element.local_name().as_ref())
                {
                    let resolver = reader.resolver().clone();
                    let anchor_id =
                        parse_word2010_anchor_id(&element, &resolver, reader.decoder())?;
                    active = Some((kind, anchor_id, depth));
                }
            },
            Event::Empty(element) => {
                if !matches!(namespace, ResolveResult::Bound(_)) && fragment_prefix.is_none() {
                    fragment_prefix = Some(
                        element
                            .name()
                            .prefix()
                            .map(|prefix| prefix.into_inner().to_vec()),
                    );
                }
                if is_word_element(&namespace, &fragment_prefix)
                    && let Some(kind) = legacy_kind(element.local_name().as_ref())
                {
                    let resolver = reader.resolver().clone();
                    let anchor_id =
                        parse_word2010_anchor_id(&element, &resolver, reader.decoder())?;
                    anchors.push(LegacyAnchor::from_parts(kind, anchor_id));
                }
            },
            Event::End(_) => {
                if active
                    .as_ref()
                    .is_some_and(|(_, _, start_depth)| *start_depth == depth)
                {
                    let (kind, anchor_id, _) = active.take().ok_or_else(|| {
                        Error::InvalidFormat("missing legacy drawing anchor".to_string())
                    })?;
                    anchors.push(LegacyAnchor::from_parts(kind, anchor_id));
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid legacy drawing XML nesting".to_string())
                })?;
            },
            Event::Eof => {
                if active.is_some() || depth != 0 {
                    return Err(Error::InvalidFormat(
                        "unterminated legacy drawing XML".to_string(),
                    ));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(anchors)
}

#[inline]
fn legacy_kind(local_name: &[u8]) -> Option<LegacyAnchorKind> {
    match local_name {
        b"object" => Some(LegacyAnchorKind::Object),
        b"pict" => Some(LegacyAnchorKind::Picture),
        _ => None,
    }
}

fn is_word_element(
    namespace: &ResolveResult<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if is_wordprocessing_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => {
            fragment_prefix
                .as_ref()
                .and_then(|prefix| prefix.as_deref())
                == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
        ResolveResult::Bound(_) => false,
    }
}

fn parse_anchor_id(element: &BytesStart<'_>) -> Result<Option<AnchorId>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let raw_name = attribute.key.as_ref();
        let local_name = raw_name
            .rsplit(|byte| *byte == b':')
            .next()
            .unwrap_or(raw_name);
        if local_name != b"anchorId" {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "DrawingML anchor has duplicate anchorId attributes".to_string(),
            ));
        }
        let text = std::str::from_utf8(&attribute.value).map_err(|error| {
            Error::InvalidFormat(format!("DrawingML anchorId is not UTF-8: {error}"))
        })?;
        value = Some(parse_anchor_id_text(text)?);
    }
    Ok(value)
}

#[inline]
fn parse_preset(element: &BytesStart<'_>) -> Result<Preset> {
    let value = strict_attribute(element, b"prst")?.ok_or_else(|| {
        Error::InvalidFormat("DrawingML prstGeom is missing required prst".to_string())
    })?;
    value.parse().map_err(|error| {
        Error::InvalidFormat(format!("invalid DrawingML shape preset '{value}': {error}"))
    })
}

fn number_attribute(element: &BytesStart<'_>, name: &[u8], default: i64) -> Result<i64> {
    Ok(inert_attribute(element, name)
        .and_then(|value| value.parse().ok())
        .unwrap_or(default))
}

/// Read a numeric or descriptive attribute as inert inventory.
///
/// The inventory historically treats numeric whitespace and DrawingML tokens
/// as authored: invalid numbers fall back to their element defaults and invalid
/// UTF-8 in descriptive metadata is ignored. Unknown/inert extension
/// attributes are never normalized or interpreted.
fn inert_attribute(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    let mut value = None;
    for attribute in element.attributes().flatten() {
        if attribute.key.as_ref() == name {
            if let Ok(decoded) = std::str::from_utf8(&attribute.value) {
                value = Some(decoded.to_owned());
            }
        }
    }
    value
}

/// Read the closed DrawingML preset token with the original strict failure
/// behavior for malformed attributes and non-UTF-8 values.
fn strict_attribute(element: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() == name {
            let decoded = std::str::from_utf8(&attribute.value).map_err(|error| {
                Error::InvalidFormat(format!("DrawingML prstGeom@prst is not UTF-8: {error}"))
            })?;
            value = Some(decoded.to_owned());
        }
    }
    Ok(value)
}
