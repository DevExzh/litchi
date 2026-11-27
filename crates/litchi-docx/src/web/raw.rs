//! Raw web-settings XML and OPC boundaries.
//!
//! The normal package API is `Package::edit_web`. Detached XML and direct OPC
//! operations remain available here with an explicit caller policy.

pub use super::limit::Limits;

use super::codec;
use super::model::{Conformance, Settings};
use super::package;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::OpcPackage;
use litchi_opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Parse detached web-settings XML under `limits`.
pub fn parse(xml: &[u8], limits: &Limits) -> Result<(Settings, Conformance)> {
    codec::parse(xml, limits)
}

/// Read a web-settings part under `limits`.
pub fn read(part: &dyn Part, limits: &Limits) -> Result<(Settings, Conformance)> {
    codec::read(part, limits)
}

/// Serialize web-settings XML under `limits`.
pub fn write(settings: &Settings, conformance: Conformance, limits: &Limits) -> Result<Vec<u8>> {
    let xml = codec::write_with_limits(settings, conformance, limits)?;
    admit_xml(&xml, limits)?;
    Ok(xml)
}

/// Put document-owned web settings under `limits`.
pub fn put(
    package: &mut OpcPackage,
    settings: Settings,
    conformance: Conformance,
    limits: &Limits,
) -> Result<bool> {
    admit_package(package, limits)?;
    let xml = codec::write_with_limits(&settings, conformance, limits)?;
    admit_xml(&xml, limits)?;
    package::put_encoded(package, settings, conformance, xml, limits)
}

/// Remove document-owned web settings under `limits`.
pub fn remove(package: &mut OpcPackage, limits: &Limits) -> Result<bool> {
    admit_package(package, limits)?;
    package::remove_with_limits(package, limits)
}

pub(crate) fn admit_xml(xml: &[u8], limits: &Limits) -> Result<()> {
    if xml.len() > limits.xml_bytes() {
        return Err(Error::WebLimit {
            kind: crate::web::LimitKind::XmlBytes,
            actual: xml.len(),
            maximum: limits.xml_bytes(),
        });
    }
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut text_bytes = 0usize;
    let mut events = 0usize;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        events = events.saturating_add(1);
        if events > limits.events() {
            return Err(Error::WebLimit {
                kind: crate::web::LimitKind::XmlEvents,
                actual: events,
                maximum: limits.events(),
            });
        }
        match event {
            Event::Start(element) => {
                depth = depth.saturating_add(1);
                nodes = nodes.saturating_add(1);
                check_shape(depth, nodes, attributes, limits)?;
                attributes = attributes.saturating_add(element.unchecked_attributes().count());
                check_shape(depth, nodes, attributes, limits)?;
            },
            Event::Empty(element) => {
                nodes = nodes.saturating_add(1);
                check_shape(depth.saturating_add(1), nodes, attributes, limits)?;
                attributes = attributes.saturating_add(element.unchecked_attributes().count());
                check_shape(depth.saturating_add(1), nodes, attributes, limits)?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("web XML has an unexpected closing element".to_owned())
                })?;
            },
            Event::Text(value) => {
                text_bytes = text_bytes.saturating_add(value.as_ref().len());
                if text_bytes > limits.text_bytes() {
                    return Err(Error::WebLimit {
                        kind: crate::web::LimitKind::TextBytes,
                        actual: text_bytes,
                        maximum: limits.text_bytes(),
                    });
                }
            },
            Event::CData(value) => {
                text_bytes = text_bytes.saturating_add(value.as_ref().len());
                if text_bytes > limits.text_bytes() {
                    return Err(Error::WebLimit {
                        kind: crate::web::LimitKind::TextBytes,
                        actual: text_bytes,
                        maximum: limits.text_bytes(),
                    });
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "web XML contains a forbidden declaration".to_owned(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::Invalid("web XML is unterminated".to_owned()));
    }
    Ok(())
}

fn check_shape(depth: usize, nodes: usize, attributes: usize, limits: &Limits) -> Result<()> {
    if depth > limits.depth() {
        return Err(Error::WebLimit {
            kind: crate::web::LimitKind::XmlDepth,
            actual: depth,
            maximum: limits.depth(),
        });
    }
    if nodes > limits.nodes() {
        return Err(Error::WebLimit {
            kind: crate::web::LimitKind::XmlNodes,
            actual: nodes,
            maximum: limits.nodes(),
        });
    }
    if attributes > limits.attributes() {
        return Err(Error::WebLimit {
            kind: crate::web::LimitKind::XmlAttributes,
            actual: attributes,
            maximum: limits.attributes(),
        });
    }
    Ok(())
}

fn admit_package(package: &OpcPackage, limits: &Limits) -> Result<()> {
    for part in package.try_iter_parts() {
        let part = part?;
        if part.content_type() == super::CONTENT_TYPE {
            admit_xml(part.blob(), limits)?;
        }
    }
    Ok(())
}
