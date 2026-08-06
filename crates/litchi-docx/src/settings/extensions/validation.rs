//! Typed and opaque settings-extension validation.

use super::super::support::{invalid, xml_error};
use super::super::{MAX_SETTINGS_XML_DEPTH, MAX_SETTINGS_XML_NODES};
use super::model::{DocumentId, Extension, MAX_EXTENSIONS, MAX_OPAQUE_BYTES};
use crate::Result;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

pub(super) fn validate_extensions(values: &[Extension]) -> Result<()> {
    if values.len() > MAX_EXTENSIONS {
        return Err(invalid(format!(
            "Word settings extension count exceeds {MAX_EXTENSIONS}"
        )));
    }

    let mut seen = [false; 6];
    for value in values {
        value.validate()?;
        let slot = match value {
            Extension::ChartTrackingRefBased(_) => Some(0),
            Extension::DocumentId(DocumentId::ParagraphContext(_)) => Some(1),
            Extension::DocumentId(DocumentId::Source(_)) => Some(2),
            Extension::ConflictMode(_) => Some(3),
            Extension::DiscardImageEditingData(_) => Some(4),
            Extension::DefaultImageDpi(_) => Some(5),
            Extension::Unknown(value) => {
                validate_opaque_xml(&value.xml)?;
                None
            },
        };
        if let Some(slot) = slot
            && std::mem::replace(&mut seen[slot], true)
        {
            return Err(invalid("duplicate typed Word settings extension"));
        }
    }
    Ok(())
}

pub(super) fn validate_opaque_xml(xml: &[u8]) -> Result<()> {
    if xml.is_empty() {
        return Err(invalid("opaque settings extension cannot be empty"));
    }
    if xml.len() > MAX_OPAQUE_BYTES {
        return Err(invalid(format!(
            "opaque settings extension exceeds {MAX_OPAQUE_BYTES} bytes"
        )));
    }
    std::str::from_utf8(xml).map_err(|_| invalid("opaque settings extension is not UTF-8"))?;

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut saw_root = false;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?
            .into_owned();
        match event {
            Event::Start(_) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("opaque settings node counter overflow"))?;
                if nodes > MAX_SETTINGS_XML_NODES {
                    return Err(invalid(format!(
                        "opaque settings extension exceeds {MAX_SETTINGS_XML_NODES} nodes"
                    )));
                }
                if depth == 0 {
                    if saw_root {
                        return Err(invalid("opaque settings extension has multiple roots"));
                    }
                    saw_root = true;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("opaque settings extension depth overflow"))?;
                if depth > MAX_SETTINGS_XML_DEPTH {
                    return Err(invalid(format!(
                        "opaque settings extension exceeds depth {MAX_SETTINGS_XML_DEPTH}"
                    )));
                }
            },
            Event::Empty(_) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("opaque settings node counter overflow"))?;
                if nodes > MAX_SETTINGS_XML_NODES {
                    return Err(invalid(format!(
                        "opaque settings extension exceeds {MAX_SETTINGS_XML_NODES} nodes"
                    )));
                }
                if depth == 0 {
                    if saw_root {
                        return Err(invalid("opaque settings extension has multiple roots"));
                    }
                    saw_root = true;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opaque settings extension has an unexpected end"))?;
            },
            Event::Text(text) => {
                let text = text
                    .decode()
                    .map_err(|error| xml_error(error.to_string()))?;
                if depth == 0 && !text.trim().is_empty() {
                    return Err(invalid(
                        "opaque settings extension has text outside its root",
                    ));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if depth > 0 => {},
            Event::Comment(_) if depth > 0 => {},
            Event::Eof => break,
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "opaque settings extension cannot contain a declaration, DTD, or processing instruction",
                ));
            },
            _ => {
                return Err(invalid(
                    "opaque settings extension has content outside its root",
                ));
            },
        }
    }

    if !saw_root || depth != 0 {
        return Err(invalid(
            "opaque settings extension is not one complete element",
        ));
    }
    Ok(())
}
