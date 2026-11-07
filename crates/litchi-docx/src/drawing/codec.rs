#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::ref_option,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
#![expect(
    clippy::unnecessary_wraps,
    reason = "the Result signature preserves a uniform fallible codec API"
)]
//! Streaming `WordprocessingML` drawing inventory decoder.

use super::model::{Anchor, AnchorId, Kind, LegacyAnchor, LegacyAnchorKind, Object};
use super::validation::{parse_anchor_id_text, parse_word2010_anchor_id};
use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use litchi_core::unit::EMUS_PER_INCH;
use litchi_drawingml::geom::Preset;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_ooxml_common::xml::attributes::first_wins;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use smallvec::SmallVec;

/// Parse drawing objects from a paragraph XML slice.
pub(crate) fn parse(xml_bytes: &[u8]) -> Result<SmallVec<[Object; 4]>> {
    let mut reader = Reader::from_reader(xml_bytes);
    reader.config_mut().trim_text(false);
    let mut objects = SmallVec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut skipped = None;
    let mut drawing_depth = None;
    let mut group_depth = None;
    let mut anchor_depth = None;
    let mut frame_extent = false;
    let mut in_shape = false;
    let mut has_shape = false;
    let mut in_text_box = false;
    let mut in_text_content = false;
    let mut in_text = false;
    let mut text_paragraphs = 0usize;
    let mut width_emu = EMUS_PER_INCH;
    let mut height_emu = EMUS_PER_INCH;
    let mut x_emu = 0;
    let mut y_emu = 0;
    let mut description = String::new();
    let mut name = String::new();
    let mut preset = None;
    let mut unknown_preset = None;
    let mut has_text_box = false;
    let mut is_inline = true;
    let mut anchor_id = None;
    let mut text_content = String::new();

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let empty = matches!(event, Event::Empty(_));
        if matches!(event, Event::Start(_)) {
            depth = depth
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("drawing nesting overflow".into()))?;
        }
        let element_depth = depth + usize::from(empty);
        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes += 1;
            if element_depth > 128 || nodes > 1_000_000 {
                return Err(Error::InvalidFormat(
                    "drawing inventory exceeds XML structure limits".into(),
                ));
            }
        }
        if let Some(skip_depth) = skipped {
            match event {
                Event::End(_) => {
                    if depth == skip_depth {
                        skipped = None;
                    }
                    depth -= 1;
                },
                Event::Eof => {
                    return Err(Error::InvalidFormat("unterminated drawing child".into()));
                },
                _ => {},
            }
            continue;
        }
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let local = element.local_name();
                let in_drawing = drawing_depth.is_some();
                if group_depth == Some(element_depth - 1)
                    && !matches!(local.as_ref(), b"grpSpPr" | b"cNvGrpSpPr")
                {
                    // Keep only the group's own properties. Child drawings
                    // are opaque to this flat inventory, even when they use
                    // familiar shape, picture or nested-group element names.
                    if !empty {
                        skipped = Some(element_depth);
                    }
                    continue;
                }
                match local.as_ref() {
                    b"drawing" if in_drawing => {
                        // A nested drawing belongs to a subsidiary text-box
                        // story; it cannot overwrite the outer anchor.
                        if !empty {
                            skipped = Some(element_depth);
                        }
                    },
                    b"drawing" => {
                        if empty {
                            continue;
                        }
                        drawing_depth = Some(element_depth);
                        group_depth = None;
                        anchor_depth = None;
                        frame_extent = false;
                        width_emu = EMUS_PER_INCH;
                        height_emu = EMUS_PER_INCH;
                        x_emu = 0;
                        y_emu = 0;
                        description.clear();
                        name.clear();
                        preset = None;
                        unknown_preset = None;
                        has_text_box = false;
                        has_shape = false;
                        in_shape = false;
                        in_text_box = false;
                        in_text_content = false;
                        in_text = false;
                        text_paragraphs = 0;
                        is_inline = true;
                        anchor_id = None;
                        text_content.clear();
                    },
                    b"inline" | b"anchor" if in_drawing => {
                        is_inline = local.as_ref() == b"inline";
                        anchor_id = parse_anchor_id(&element)?;
                        anchor_depth = Some(element_depth);
                    },
                    b"extent" if in_drawing && anchor_depth == Some(element_depth - 1) => {
                        width_emu = number_attribute(&element, b"cx", EMUS_PER_INCH)?;
                        height_emu = number_attribute(&element, b"cy", EMUS_PER_INCH)?;
                        frame_extent = true;
                    },
                    b"wgp" | b"grpSp" if in_drawing => {
                        // The flat inventory has no group-child model. Keep
                        // the outer drawing as Other, without mistaking its
                        // last child for the group's geometry or text.
                        if !empty {
                            group_depth = Some(element_depth);
                        }
                    },
                    b"off" if in_drawing => {
                        x_emu = number_attribute(&element, b"x", 0)?;
                        y_emu = number_attribute(&element, b"y", 0)?;
                    },
                    b"ext" if in_drawing && !frame_extent => {
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
                        in_shape = !empty;
                        has_shape = true;
                    },
                    b"prstGeom" if in_shape => match parse_preset(&element)? {
                        Ok(value) => {
                            preset = Some(value);
                            unknown_preset = None;
                        },
                        Err(token) => {
                            preset = None;
                            unknown_preset = Some(token);
                        },
                    },
                    b"txbx" if in_shape => {
                        in_text_box = !empty;
                        has_text_box = true;
                    },
                    b"txbxContent" if in_text_box => {
                        in_text_content = !empty;
                    },
                    b"p" if in_text_content => {
                        if text_paragraphs > 0 {
                            text_content.push('\n');
                        }
                        text_paragraphs += 1;
                    },
                    b"t" if in_text_content => {
                        in_text = !empty;
                    },
                    b"tab" if in_text_content => text_content.push('\t'),
                    b"br" | b"cr" if in_text_content => text_content.push('\n'),
                    _ => {},
                }
            },
            Event::Text(text) if in_text => {
                let decoded = text
                    .xml_content(quick_xml::XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let value = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                text_content.push_str(&value);
            },
            Event::CData(text) if in_text => {
                let value = text
                    .xml_content(quick_xml::XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                text_content.push_str(&value);
            },
            Event::GeneralRef(reference) if in_text => {
                let value = litchi_ooxml_common::xml::decode_xml_reference(&reference)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                text_content.push_str(&value);
            },
            Event::End(element) => {
                if group_depth == Some(depth) {
                    group_depth = None;
                }
                if drawing_depth == Some(depth) {
                    drawing_depth = None;
                    let kind = if has_text_box {
                        Kind::TextBox
                    } else if has_shape {
                        Kind::Shape
                    } else {
                        Kind::Other
                    };
                    objects.push(
                        Object::from_inventory(
                            std::mem::take(&mut name),
                            std::mem::take(&mut description),
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
                            std::mem::take(&mut text_content),
                        )
                        .with_unknown_preset(unknown_preset.take()),
                    );
                } else {
                    match element.local_name().as_ref() {
                        b"wsp" => in_shape = false,
                        b"txbx" => in_text_box = false,
                        b"txbxContent" => {
                            in_text_content = false;
                            in_text = false;
                        },
                        b"t" => in_text = false,
                        _ => {},
                    }
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid drawing XML nesting".into()))?;
            },
            Event::Eof => {
                if depth != 0 || drawing_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "unterminated drawing inventory".into(),
                    ));
                }
                break;
            },
            _ => {},
        }
    }
    Ok(objects)
}

/// Parse legacy Word `w:object` and `w:pict` anchors from a paragraph XML
/// slice. Only the checked `anchorId` metadata is projected; child VML, OLE,
/// and image content remains inert.
#[cfg(test)]
pub(crate) fn parse_legacy(xml_bytes: &[u8]) -> Result<SmallVec<[LegacyAnchor; 4]>> {
    parse_legacy_with_context(xml_bytes, &std::sync::Arc::from([]))
}

pub(crate) fn parse_legacy_with_context(
    xml_bytes: &[u8],
    namespaces: &crate::namespace::NamespaceBindings,
) -> Result<SmallVec<[LegacyAnchor; 4]>> {
    let mut reader = NsReader::from_reader(xml_bytes);
    for (prefix, namespace) in namespaces.iter() {
        let prefix = prefix
            .as_deref()
            .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|_| Error::Xml("invalid inherited legacy-anchor namespace".into()))?;
    }
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
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
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
    for attribute in element.checked_attributes() {
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

/// Read `prstGeom@prst`: the closed preset, or the raw token when it names
/// no `ST_ShapeType` value (the empty string when `prst` is missing).
///
/// LibreOffice imports such a shape with an unknown preset type
/// (`PresetShapeGeometryContext`) and Word renders it best-effort, so the
/// object is kept rather than failing the paragraph's inventory
/// (LENIENCY_AUDIT 2.10). Malformed attributes stay errors.
#[inline]
fn parse_preset(element: &BytesStart<'_>) -> Result<std::result::Result<Preset, String>> {
    let Some(value) = strict_attribute(element, b"prst")? else {
        return Ok(Err(String::new()));
    };
    Ok(value.parse().map_err(|_unknown| value))
}

fn number_attribute(element: &BytesStart<'_>, name: &[u8], default: i64) -> Result<i64> {
    Ok(inert_attribute(element, name)
        .and_then(|value| value.parse().ok())
        .unwrap_or(default))
}

/// Read a numeric or descriptive attribute as inert inventory.
///
/// The inventory historically treats numeric whitespace and `DrawingML` tokens
/// as authored: invalid numbers fall back to their element defaults and invalid
/// UTF-8 in descriptive metadata is ignored. Unknown/inert extension
/// attributes are never normalized or interpreted. A repeated attribute name
/// keeps its first occurrence.
fn inert_attribute(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    let mut value = None;
    for attribute in first_wins(element).flatten() {
        if attribute.key.as_ref() == name
            && let Ok(decoded) = std::str::from_utf8(&attribute.value)
        {
            value = Some(decoded.to_owned());
        }
    }
    value
}

/// Read the closed `DrawingML` preset token with the original strict failure
/// behavior for malformed attributes and non-UTF-8 values.
fn strict_attribute(element: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.checked_attributes() {
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
