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
use litchi_opc::SourceLimits;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Load one owner resource under `limits`.
pub fn load(package: &OpcPackage, owner: Owner, limits: &Limits) -> Result<Snapshot> {
    admit_package(package, owner, limits)?;
    super::load_with_limits(package, owner, source_limits(limits)?)
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
    super::put_with_limits(package, owner, resource, source_limits(limits)?)
}

/// Remove one owner resource under `limits`.
pub fn remove(package: &mut OpcPackage, owner: Owner, limits: &Limits) -> Result<bool> {
    admit_package(package, owner, limits)?;
    super::remove_with_limits(package, owner, source_limits(limits)?)
}

fn source_limits(limits: &Limits) -> Result<SourceLimits> {
    let xml_bytes = u64::try_from(limits.xml_bytes())
        .map_err(|_| Error::InvalidFormat("effects XML byte limit overflows u64".to_owned()))?;
    SourceLimits::builder()
        .max_part_bytes(xml_bytes)
        .and_then(|value| value.max_events(limits.events()))
        .and_then(|value| value.max_depth(limits.depth()))
        .and_then(|value| value.max_xml_attribute_bytes(limits.attribute_bytes()))
        .and_then(|value| value.build())
        .map_err(Error::Opc)
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
    // Preflight the selected owner bytes before inspect_graph parses them
    // under the package-wide source policy.  The graph validator still owns
    // topology, content-type, and semantic checks; this pass only establishes
    // the caller's tighter effects budget before any of those allocations.
    let source_names = owner_sources(package, owner)?;
    for source_name in source_names {
        let source = package.get_part(&source_name)?;
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == super::RELATIONSHIP_TYPE)
        {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            let target = package.get_part(&target)?;
            if target.content_type() == super::CONTENT_TYPE {
                admit_xml(target.blob(), limits)?;
            }
        }
    }
    Ok(())
}

fn owner_sources(package: &OpcPackage, owner: Owner) -> Result<Vec<litchi_opc::PackURI>> {
    let main = package.main_document_part()?.partname().clone();
    if owner == Owner::MainDocument {
        return Ok(vec![main]);
    }

    let main_part = package.get_part(&main)?;
    let mut sources = Vec::new();
    for relationship in main_part.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            super::GLOSSARY_RELATIONSHIP_TYPE | super::STRICT_GLOSSARY_RELATIONSHIP_TYPE
        )
    }) {
        if relationship.is_external() {
            continue;
        }
        sources.push(relationship.target_partname()?);
    }
    Ok(sources)
}
