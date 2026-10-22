//! Raw stylesWithEffects XML and OPC graph boundary.
//!
//! Package callers normally use `Package::edit_styles_with_effects`. Direct
//! XML construction and graph mutation remain available here with an
//! explicit caller-selected limit policy.

pub use super::limit::Limits;

use super::{Owner, Resource, Snapshot};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::OpcPackage;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Load one owner resource under `limits`.
pub fn load(package: &OpcPackage, owner: Owner, limits: &Limits) -> Result<Snapshot> {
    admit_package(package, owner, limits)?;
    super::load(package, owner)
}

/// Put one owner resource under `limits`.
pub fn put(
    package: &mut OpcPackage,
    owner: Owner,
    resource: Resource,
    limits: &Limits,
) -> Result<bool> {
    admit_package(package, owner, limits)?;
    admit_xml(resource.xml_bytes(), limits)?;
    super::put(package, owner, resource)
}

/// Remove one owner resource under `limits`.
pub fn remove(package: &mut OpcPackage, owner: Owner, limits: &Limits) -> Result<bool> {
    admit_package(package, owner, limits)?;
    super::remove(package, owner)
}

pub(crate) fn admit_xml(xml: &[u8], limits: &Limits) -> Result<()> {
    if xml.len() > limits.xml_bytes() {
        return Err(Error::StylesWithEffectsLimit {
            kind: crate::styles::effects::LimitKind::XmlBytes,
            actual: xml.len(),
            maximum: limits.xml_bytes(),
        });
    }
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut events = 0usize;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut attribute_bytes = 0usize;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        events = events.saturating_add(1);
        if events > limits.events() {
            return Err(Error::StylesWithEffectsLimit {
                kind: crate::styles::effects::LimitKind::XmlEvents,
                actual: events,
                maximum: limits.events(),
            });
        }
        match event {
            Event::Start(element) => {
                depth = depth.saturating_add(1);
                nodes = nodes.saturating_add(1);
                check_shape(depth, nodes, attributes, limits)?;
                check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
            },
            Event::Empty(element) => {
                nodes = nodes.saturating_add(1);
                check_shape(depth.saturating_add(1), nodes, attributes, limits)?;
                check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid(
                        "stylesWithEffects XML has an unexpected closing element".to_owned(),
                    )
                })?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "stylesWithEffects XML contains a forbidden declaration".to_owned(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::Invalid(
            "stylesWithEffects XML is unterminated".to_owned(),
        ));
    }
    Ok(())
}

fn check_shape(depth: usize, nodes: usize, attributes: usize, limits: &Limits) -> Result<()> {
    if depth > limits.depth() {
        return Err(Error::StylesWithEffectsLimit {
            kind: crate::styles::effects::LimitKind::XmlDepth,
            actual: depth,
            maximum: limits.depth(),
        });
    }
    if nodes > limits.nodes() {
        return Err(Error::StylesWithEffectsLimit {
            kind: crate::styles::effects::LimitKind::XmlNodes,
            actual: nodes,
            maximum: limits.nodes(),
        });
    }
    if attributes > limits.attributes() {
        return Err(Error::StylesWithEffectsLimit {
            kind: crate::styles::effects::LimitKind::XmlAttributes,
            actual: attributes,
            maximum: limits.attributes(),
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
            .checked_add(
                attribute
                    .key
                    .as_ref()
                    .len()
                    .saturating_add(attribute.value.len()),
            )
            .ok_or(Error::StylesWithEffectsLimit {
                kind: crate::styles::effects::LimitKind::XmlAttributeBytes,
                actual: usize::MAX,
                maximum: limits.attribute_bytes(),
            })?;
        check_shape(0, 0, *attributes, limits)?;
        if *attribute_bytes > limits.attribute_bytes() {
            return Err(Error::StylesWithEffectsLimit {
                kind: crate::styles::effects::LimitKind::XmlAttributeBytes,
                actual: *attribute_bytes,
                maximum: limits.attribute_bytes(),
            });
        }
    }
    Ok(())
}

fn admit_package(package: &OpcPackage, owner: Owner, limits: &Limits) -> Result<()> {
    let graph = super::inspect_graph(package)?;
    let Some(binding) = graph.binding(owner) else {
        return Ok(());
    };
    let part = package.get_part(&binding.target)?;
    admit_xml(part.blob(), limits)
}
