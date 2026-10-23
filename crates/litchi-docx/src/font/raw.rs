//! Raw font-table XML and OPC graph codecs.
//!
//! The typed owner family is the normal package editing surface. This module
//! retains the lower-level table and relationship-graph capability for callers
//! that intentionally operate on an [`OpcPackage`], with an explicit policy
//! admitted before the codec retains input state.

pub use super::limit::Limits;
pub use super::model::raw::Attr;

use super::codec;
use super::model::{Conformance, FONT_CT, FT_CT, Table, invalid};
use super::package;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::OpcPackage;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashSet;

/// Parse one detached font table under `limits`.
pub fn parse(xml: &[u8], limits: &Limits) -> Result<Table> {
    codec::parse(xml, limits)
}

/// Serialize one detached font table under `limits`.
pub fn write(table: &Table, conformance: Conformance, limits: &Limits) -> Result<Vec<u8>> {
    admit_table(table, conformance, limits)?;
    let estimate = estimated_xml_bytes(table, conformance)?;
    if estimate > limits.xml_bytes() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::XmlBytes,
            actual: estimate,
            maximum: limits.xml_bytes(),
        });
    }
    let xml = codec::write(table, conformance)?;
    admit_xml(&xml, limits)?;
    Ok(xml)
}

/// Read the document-owned font table and embedded resources under `limits`.
pub fn read(package: &OpcPackage, limits: &Limits) -> Result<Option<Table>> {
    admit_package(package, limits)?;
    package::read_with_limits(package, limits)
}

/// Put a complete font table and its embedded resources under `limits`.
pub fn put(
    package: &mut OpcPackage,
    table: Table,
    conformance: Conformance,
    limits: &Limits,
) -> Result<bool> {
    admit_table(&table, conformance, limits)?;
    admit_package(package, limits)?;
    package::put_checked(package, table, conformance, limits)
}

/// Remove the document-owned font graph under `limits`.
pub fn remove(package: &mut OpcPackage, limits: &Limits) -> Result<bool> {
    admit_package(package, limits)?;
    package::remove_with_limits(package, limits)
}

fn admit_package(package: &OpcPackage, limits: &Limits) -> Result<()> {
    let mut font_parts = 0usize;
    let mut total_font_bytes = 0usize;
    for part in package.try_iter_parts() {
        let part = part?;
        if part.content_type() == FT_CT {
            admit_xml(part.blob(), limits)?;
        } else if part.content_type() == FONT_CT {
            font_parts = font_parts.saturating_add(1);
            if font_parts > limits.fonts() {
                return Err(Error::FontLimit {
                    kind: crate::font::LimitKind::FontCount,
                    actual: font_parts,
                    maximum: limits.fonts(),
                });
            }
            if part.blob().len() > limits.embedded_font_bytes() {
                return Err(Error::FontLimit {
                    kind: crate::font::LimitKind::EmbeddedFontBytes,
                    actual: part.blob().len(),
                    maximum: limits.embedded_font_bytes(),
                });
            }
            total_font_bytes =
                total_font_bytes
                    .checked_add(part.blob().len())
                    .ok_or(Error::FontLimit {
                        kind: crate::font::LimitKind::AggregateBytes,
                        actual: usize::MAX,
                        maximum: limits.aggregate_font_bytes(),
                    })?;
            if total_font_bytes > limits.aggregate_font_bytes() {
                return Err(Error::FontLimit {
                    kind: crate::font::LimitKind::AggregateBytes,
                    actual: total_font_bytes,
                    maximum: limits.aggregate_font_bytes(),
                });
            }
        }
    }
    Ok(())
}

pub(crate) fn admit_table(table: &Table, conformance: Conformance, limits: &Limits) -> Result<()> {
    if table.fonts.len() > limits.fonts() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::FontCount,
            actual: table.fonts.len(),
            maximum: limits.fonts(),
        });
    }
    let mut total = 0usize;
    let mut attribute_count = 0usize;
    let mut attribute_bytes = 0usize;
    admit_pair(
        "xmlns:w",
        conformance.word(),
        &mut attribute_count,
        &mut attribute_bytes,
        limits,
    )?;
    admit_pair(
        "xmlns:r",
        conformance.rel(),
        &mut attribute_count,
        &mut attribute_bytes,
        limits,
    )?;
    for attribute in table.namespaces.iter().chain(&table.extension_attributes) {
        admit_attribute(
            attribute,
            &mut attribute_count,
            &mut attribute_bytes,
            limits,
        )?;
    }
    let mut resources = HashSet::new();
    for font in &table.fonts {
        for value in [Some(font.name.as_str()), font.alternate_name.as_deref()]
            .into_iter()
            .flatten()
        {
            if value.len() > limits.text_bytes() {
                return Err(Error::FontLimit {
                    kind: crate::font::LimitKind::Text,
                    actual: value.len(),
                    maximum: limits.text_bytes(),
                });
            }
        }
        for attribute in &font.extension_attributes {
            admit_attribute(
                attribute,
                &mut attribute_count,
                &mut attribute_bytes,
                limits,
            )?;
        }
        admit_pair(
            "w:name",
            &font.name,
            &mut attribute_count,
            &mut attribute_bytes,
            limits,
        )?;
        if let Some(value) = &font.alternate_name {
            admit_pair(
                "w:val",
                value,
                &mut attribute_count,
                &mut attribute_bytes,
                limits,
            )?;
        }
        if font.panose.is_some() {
            admit_pair(
                "w:val",
                "00000000000000000000",
                &mut attribute_count,
                &mut attribute_bytes,
                limits,
            )?;
        }
        if let Some(value) = font.character_set {
            let value = if conformance == Conformance::Transitional {
                "00"
            } else {
                value.strict_name().unwrap_or("")
            };
            admit_pair(
                if conformance == Conformance::Transitional {
                    "w:val"
                } else {
                    "w:characterSet"
                },
                value,
                &mut attribute_count,
                &mut attribute_bytes,
                limits,
            )?;
        }
        if let Some(value) = font.family {
            admit_pair(
                "w:val",
                value.text(),
                &mut attribute_count,
                &mut attribute_bytes,
                limits,
            )?;
        }
        if font.not_true_type == Some(false) {
            admit_pair(
                "w:val",
                "0",
                &mut attribute_count,
                &mut attribute_bytes,
                limits,
            )?;
        }
        if let Some(value) = font.pitch {
            admit_pair(
                "w:val",
                value.text(),
                &mut attribute_count,
                &mut attribute_bytes,
                limits,
            )?;
        }
        if font.signature.is_some() {
            for name in ["w:usb0", "w:usb1", "w:usb2", "w:usb3", "w:csb0", "w:csb1"] {
                admit_pair(
                    name,
                    "00000000",
                    &mut attribute_count,
                    &mut attribute_bytes,
                    limits,
                )?;
            }
        }
        for embedded in &font.embedded_fonts {
            if embedded.relationship_id.len() > limits.text_bytes() {
                return Err(Error::FontLimit {
                    kind: crate::font::LimitKind::Text,
                    actual: embedded.relationship_id.len(),
                    maximum: limits.text_bytes(),
                });
            }
            for attribute in &embedded.extension_attributes {
                admit_attribute(
                    attribute,
                    &mut attribute_count,
                    &mut attribute_bytes,
                    limits,
                )?;
            }
            admit_pair(
                "r:id",
                &embedded.relationship_id,
                &mut attribute_count,
                &mut attribute_bytes,
                limits,
            )?;
            if embedded.font_key.is_some() {
                admit_pair(
                    "w:fontKey",
                    "{00000000-0000-0000-0000-000000000000}",
                    &mut attribute_count,
                    &mut attribute_bytes,
                    limits,
                )?;
            }
            if let Some(value) = embedded.subsetted {
                admit_pair(
                    "w:subsetted",
                    if value { "1" } else { "0" },
                    &mut attribute_count,
                    &mut attribute_bytes,
                    limits,
                )?;
            }
            if let Some(resource) = &embedded.resource {
                if resource.data.len() > limits.embedded_font_bytes() {
                    return Err(Error::FontLimit {
                        kind: crate::font::LimitKind::EmbeddedFontBytes,
                        actual: resource.data.len(),
                        maximum: limits.embedded_font_bytes(),
                    });
                }
                if resource.part_name.is_empty() || resources.insert(resource.part_name.as_str()) {
                    total = total
                        .checked_add(resource.data.len())
                        .ok_or(Error::FontLimit {
                            kind: crate::font::LimitKind::AggregateBytes,
                            actual: usize::MAX,
                            maximum: limits.aggregate_font_bytes(),
                        })?;
                }
            }
        }
    }
    if total > limits.aggregate_font_bytes() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::AggregateBytes,
            actual: total,
            maximum: limits.aggregate_font_bytes(),
        });
    }
    Ok(())
}

fn admit_attribute(
    attribute: &Attr,
    count: &mut usize,
    bytes: &mut usize,
    limits: &Limits,
) -> Result<()> {
    admit_pair(
        attribute.qualified_name(),
        attribute.value(),
        count,
        bytes,
        limits,
    )
}

fn admit_pair(
    name: &str,
    value: &str,
    count: &mut usize,
    bytes: &mut usize,
    limits: &Limits,
) -> Result<()> {
    for value in [name, value] {
        if value.len() > limits.text_bytes() {
            return Err(Error::FontLimit {
                kind: crate::font::LimitKind::Text,
                actual: value.len(),
                maximum: limits.text_bytes(),
            });
        }
    }
    *count = count.checked_add(1).ok_or(Error::FontLimit {
        kind: crate::font::LimitKind::XmlAttributeCount,
        actual: usize::MAX,
        maximum: limits.attributes(),
    })?;
    *bytes = bytes
        .checked_add(name.len())
        .and_then(|total| total.checked_add(value.len()))
        .ok_or(Error::FontLimit {
            kind: crate::font::LimitKind::XmlAttributeBytes,
            actual: usize::MAX,
            maximum: limits.attribute_bytes(),
        })?;
    if *count > limits.attributes() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::XmlAttributeCount,
            actual: *count,
            maximum: limits.attributes(),
        });
    }
    if *bytes > limits.attribute_bytes() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::XmlAttributeBytes,
            actual: *bytes,
            maximum: limits.attribute_bytes(),
        });
    }
    Ok(())
}

fn estimated_xml_bytes(table: &Table, conformance: Conformance) -> Result<usize> {
    let mut bytes = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.len();
    bytes = add_len(bytes, b"<w:fonts xmlns:w=\"")?;
    bytes = add_len(bytes, conformance.word().as_bytes())?;
    bytes = add_len(bytes, b"\" xmlns:r=\"")?;
    bytes = add_len(bytes, conformance.rel().as_bytes())?;
    bytes = add_len(bytes, b"\"")?;
    for attribute in table.namespaces.iter().chain(&table.extension_attributes) {
        bytes = add_attribute_len(bytes, attribute.qualified_name(), attribute.value())?;
    }
    if table.fonts.is_empty() {
        return add_len(bytes, b"/>");
    }
    bytes = add_len(bytes, b">")?;
    for font in &table.fonts {
        bytes = add_len(bytes, b"<w:font")?;
        for attribute in &font.extension_attributes {
            bytes = add_attribute_len(bytes, attribute.qualified_name(), attribute.value())?;
        }
        bytes = add_attribute_len(bytes, "w:name", &font.name)?;
        let empty = font.alternate_name.is_none()
            && font.panose.is_none()
            && font.character_set.is_none()
            && font.family.is_none()
            && font.not_true_type.is_none()
            && font.pitch.is_none()
            && font.signature.is_none()
            && font.embedded_fonts.is_empty();
        if empty {
            bytes = add_len(bytes, b"/>")?;
            continue;
        }
        bytes = add_len(bytes, b">")?;
        if let Some(value) = &font.alternate_name {
            bytes = add_value_leaf_len(bytes, "altName", value)?;
        }
        if font.panose.is_some() {
            bytes = add_value_leaf_len(bytes, "panose1", "00000000000000000000")?;
        }
        if let Some(value) = font.character_set {
            bytes = add_len(bytes, b"<w:charset")?;
            let value = match conformance {
                Conformance::Transitional => "00",
                Conformance::Strict => value.strict_name().unwrap_or(""),
            };
            if value.is_empty() {
                return Err(invalid("legacy charset has no Strict representation"));
            }
            bytes = add_attribute_len(
                bytes,
                if conformance == Conformance::Transitional {
                    "w:val"
                } else {
                    "w:characterSet"
                },
                value,
            )?;
            bytes = add_len(bytes, b"/>")?;
        }
        if let Some(value) = font.family {
            bytes = add_value_leaf_len(bytes, "family", value.text())?;
        }
        if let Some(value) = font.not_true_type {
            bytes = add_len(bytes, b"<w:notTrueType")?;
            if !value {
                bytes = add_attribute_len(bytes, "w:val", "0")?;
            }
            bytes = add_len(bytes, b"/>")?;
        }
        if let Some(value) = font.pitch {
            bytes = add_value_leaf_len(bytes, "pitch", value.text())?;
        }
        if font.signature.is_some() {
            bytes = add_len(bytes, b"<w:sig")?;
            for name in ["w:usb0", "w:usb1", "w:usb2", "w:usb3", "w:csb0", "w:csb1"] {
                bytes = add_attribute_len(bytes, name, "00000000")?;
            }
            bytes = add_len(bytes, b"/>")?;
        }
        for embed in &font.embedded_fonts {
            bytes = add_len(bytes, b"<w:")?;
            bytes = add_len(bytes, embed.style.element().as_bytes())?;
            for attribute in &embed.extension_attributes {
                bytes = add_attribute_len(bytes, attribute.qualified_name(), attribute.value())?;
            }
            bytes = add_attribute_len(bytes, "r:id", &embed.relationship_id)?;
            if embed.font_key.is_some() {
                bytes = add_attribute_len(
                    bytes,
                    "w:fontKey",
                    "{00000000-0000-0000-0000-000000000000}",
                )?;
            }
            if let Some(value) = embed.subsetted {
                bytes = add_attribute_len(bytes, "w:subsetted", if value { "1" } else { "0" })?;
            }
            bytes = add_len(bytes, b"/>")?;
        }
        bytes = add_len(bytes, b"</w:font>")?;
    }
    add_len(bytes, b"</w:fonts>")
}

fn add_len(current: usize, value: &[u8]) -> Result<usize> {
    current
        .checked_add(value.len())
        .ok_or_else(|| invalid("font XML size overflow"))
}

fn add_attribute_len(current: usize, name: &str, value: &str) -> Result<usize> {
    let escaped = escaped_len(value);
    current
        .checked_add(1)
        .and_then(|value| value.checked_add(name.len()))
        .and_then(|value| value.checked_add(2))
        .and_then(|value| value.checked_add(escaped))
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| invalid("font XML size overflow"))
}

fn add_value_leaf_len(current: usize, name: &str, value: &str) -> Result<usize> {
    let current = add_len(current, b"<w:")?;
    let current = add_len(current, name.as_bytes())?;
    let current = add_attribute_len(current, "w:val", value)?;
    add_len(current, b"/>")
}

fn escaped_len(value: &str) -> usize {
    value
        .chars()
        .map(|value| match value {
            '&' => 5,
            '<' => 4,
            '"' => 6,
            '\t' | '\n' | '\r' => 5,
            value => value.len_utf8(),
        })
        .sum()
}

pub(crate) fn admit_xml(xml: &[u8], limits: &Limits) -> Result<()> {
    if xml.len() > limits.xml_bytes() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::XmlBytes,
            actual: xml.len(),
            maximum: limits.xml_bytes(),
        });
    }
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut attribute_bytes = 0usize;
    let mut text_bytes = 0usize;
    let mut font_count = 0usize;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) => {
                check_font_count(&element, &mut font_count, limits)?;
                depth = depth.saturating_add(1);
                nodes = nodes.saturating_add(1);
                check_shape(depth, nodes, limits)?;
                check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
            },
            Event::Empty(element) => {
                check_font_count(&element, &mut font_count, limits)?;
                nodes = nodes.saturating_add(1);
                check_shape(depth.saturating_add(1), nodes, limits)?;
                check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("font XML has an unexpected closing element".to_owned())
                })?;
            },
            Event::Text(value) => {
                check_text_bytes(value.as_ref().len(), &mut text_bytes, limits)?;
            },
            Event::CData(value) => {
                check_text_bytes(value.as_ref().len(), &mut text_bytes, limits)?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "font XML contains a forbidden declaration".to_owned(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::Invalid("font XML is unterminated".to_owned()));
    }
    Ok(())
}

fn check_font_count(element: &BytesStart<'_>, count: &mut usize, limits: &Limits) -> Result<()> {
    let element_name = element.name();
    let qualified = std::str::from_utf8(element_name.as_ref())
        .map_err(|error| Error::Xml(error.to_string()))?;
    let local = qualified
        .rsplit_once(':')
        .map_or(qualified, |(_, local)| local);
    if local == "font" {
        *count = count.checked_add(1).ok_or(Error::FontLimit {
            kind: crate::font::LimitKind::FontCount,
            actual: usize::MAX,
            maximum: limits.fonts(),
        })?;
        if *count > limits.fonts() {
            return Err(Error::FontLimit {
                kind: crate::font::LimitKind::FontCount,
                actual: *count,
                maximum: limits.fonts(),
            });
        }
    }
    Ok(())
}

fn check_text_bytes(value: usize, total: &mut usize, limits: &Limits) -> Result<()> {
    *total = total.checked_add(value).ok_or(Error::FontLimit {
        kind: crate::font::LimitKind::Text,
        actual: usize::MAX,
        maximum: limits.text_bytes(),
    })?;
    if *total > limits.text_bytes() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::Text,
            actual: *total,
            maximum: limits.text_bytes(),
        });
    }
    Ok(())
}

fn check_shape(depth: usize, nodes: usize, limits: &Limits) -> Result<()> {
    if depth > limits.depth() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::XmlStructure,
            actual: depth,
            maximum: limits.depth(),
        });
    }
    if nodes > limits.nodes() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::XmlNode,
            actual: nodes,
            maximum: limits.nodes(),
        });
    }
    Ok(())
}

fn check_attributes(
    element: &BytesStart<'_>,
    attributes: &mut usize,
    attribute_bytes: &mut usize,
    limits: &Limits,
) -> Result<()> {
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        *attributes = attributes.saturating_add(1);
        *attribute_bytes = attribute_bytes
            .checked_add(attribute.key.as_ref().len() + attribute.value.len())
            .ok_or(Error::FontLimit {
                kind: crate::font::LimitKind::XmlAttributeBytes,
                actual: usize::MAX,
                maximum: limits.attribute_bytes(),
            })?;
        if *attributes > limits.attributes() {
            return Err(Error::FontLimit {
                kind: crate::font::LimitKind::XmlAttributeCount,
                actual: *attributes,
                maximum: limits.attributes(),
            });
        }
        if *attribute_bytes > limits.attribute_bytes() {
            return Err(Error::FontLimit {
                kind: crate::font::LimitKind::XmlAttributeBytes,
                actual: *attribute_bytes,
                maximum: limits.attribute_bytes(),
            });
        }
    }
    Ok(())
}
