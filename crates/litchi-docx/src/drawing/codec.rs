//! Streaming WordprocessingML drawing inventory decoder.

use super::model::{Anchor, Kind, Object};
use crate::error::{Error, Result};
use litchi_core::unit::EMUS_PER_INCH;
use litchi_drawingml::geom::Preset;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
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
                        text_content.clear();
                    },
                    b"inline" if in_drawing => {
                        is_inline = true;
                    },
                    b"anchor" if in_drawing => {
                        is_inline = false;
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
