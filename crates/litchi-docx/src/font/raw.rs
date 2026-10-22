//! Raw font-table XML and OPC graph codecs.
//!
//! The typed owner family is the normal package editing surface. This module
//! retains the lower-level table and relationship-graph capability for callers
//! that intentionally operate on an [`OpcPackage`], with an explicit policy
//! admitted before the codec retains input state.

pub use super::limit::Limits;
pub use super::model::raw::Attr;

use super::codec;
use super::model::{Conformance, FONT_CT, FT_CT, Table};
use super::package;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::OpcPackage;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Parse one detached font table under `limits`.
pub fn parse(xml: &[u8], limits: &Limits) -> Result<Table> {
    admit_xml(xml, limits)?;
    codec::parse(xml)
}

/// Serialize one detached font table under `limits`.
pub fn write(table: &Table, conformance: Conformance, limits: &Limits) -> Result<Vec<u8>> {
    admit_table(table, limits)?;
    let xml = codec::write(table, conformance)?;
    admit_xml(&xml, limits)?;
    Ok(xml)
}

/// Read the document-owned font table and embedded resources under `limits`.
pub fn read(package: &OpcPackage, limits: &Limits) -> Result<Option<Table>> {
    admit_package(package, limits)?;
    package::read(package)
}

/// Put a complete font table and its embedded resources under `limits`.
pub fn put(
    package: &mut OpcPackage,
    table: Table,
    conformance: Conformance,
    limits: &Limits,
) -> Result<bool> {
    admit_table(&table, limits)?;
    admit_package(package, limits)?;
    package::put_checked(package, table, conformance, limits)
}

/// Remove the document-owned font graph under `limits`.
pub fn remove(package: &mut OpcPackage, limits: &Limits) -> Result<bool> {
    admit_package(package, limits)?;
    package::remove(package)
}

fn admit_package(package: &OpcPackage, limits: &Limits) -> Result<()> {
    for part in package.iter_parts() {
        if part.content_type() == FT_CT {
            admit_xml(part.blob(), limits)?;
        } else if part.content_type() == FONT_CT {
            if part.blob().len() > limits.embedded_font_bytes() {
                return Err(Error::FontLimit {
                    kind: crate::font::LimitKind::EmbeddedFontBytes,
                    actual: part.blob().len(),
                    maximum: limits.embedded_font_bytes(),
                });
            }
        }
    }
    Ok(())
}

fn admit_table(table: &Table, limits: &Limits) -> Result<()> {
    if table.fonts.len() > limits.fonts() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::FontCount,
            actual: table.fonts.len(),
            maximum: limits.fonts(),
        });
    }
    let mut total = 0usize;
    for font in &table.fonts {
        for value in [Some(font.name.as_str()), font.alternate_name.as_deref()] {
            if let Some(value) = value {
                if value.len() > limits.text_bytes() {
                    return Err(Error::FontLimit {
                        kind: crate::font::LimitKind::Text,
                        actual: value.len(),
                        maximum: limits.text_bytes(),
                    });
                }
            }
        }
        for embedded in &font.embedded_fonts {
            if let Some(resource) = &embedded.resource {
                if resource.data.len() > limits.embedded_font_bytes() {
                    return Err(Error::FontLimit {
                        kind: crate::font::LimitKind::EmbeddedFontBytes,
                        actual: resource.data.len(),
                        maximum: limits.embedded_font_bytes(),
                    });
                }
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
    if total > limits.aggregate_font_bytes() {
        return Err(Error::FontLimit {
            kind: crate::font::LimitKind::AggregateBytes,
            actual: total,
            maximum: limits.aggregate_font_bytes(),
        });
    }
    Ok(())
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
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) => {
                depth = depth.saturating_add(1);
                nodes = nodes.saturating_add(1);
                check_shape(depth, nodes, limits)?;
                check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
            },
            Event::Empty(element) => {
                nodes = nodes.saturating_add(1);
                check_shape(depth.saturating_add(1), nodes, limits)?;
                check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("font XML has an unexpected closing element".to_owned())
                })?;
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
    element: &quick_xml::events::BytesStart<'_>,
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
                maximum: limits.text_bytes(),
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
