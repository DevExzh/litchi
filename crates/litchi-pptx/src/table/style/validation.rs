//! Package-graph topology and semantic model validation for table styles.

use super::codec::{Parsed, parse_owned, scan, semantic_xml_eq, xml_char};
use super::model::{Conformance, Def, Link, List, Parts};
use super::{
    MAX_ATTRIBUTE_BYTES, MAX_DEPTH, MAX_GRAPH_PARTS, MAX_GRAPH_RELATIONSHIPS, MAX_NODES,
    MAX_PRESENTATION_BYTES, MAX_STYLES, P, PART_NAME_ATTEMPTS, PS, RELATIONSHIP_ID_ATTEMPTS,
    STRICT_REL, allocation, invalid, limit, xml_error,
};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

/// Borrow the optional, fully validated table-style relationship.
///
/// This is the low-level package-writer seam. It validates the complete graph
/// and returns exact producer relationship fields without copying catalog XML.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn link(package: &OpcPackage) -> Result<Option<Link<'_>>> {
    let graph = inspect_graph(package)?;
    let Some(attachment) = graph.attachment else {
        return Ok(None);
    };
    let relationship = package
        .get_part(&graph.presentation)?
        .rels()
        .get(&attachment.relationship_id)
        .ok_or_else(|| invalid("validated table-style relationship disappeared"))?;
    Ok(Some(Link {
        id: relationship.r_id(),
        kind: relationship.reltype(),
        target: relationship.target_ref(),
    }))
}

/// Validate the package graph and return its presentation conformance.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn conformance(package: &OpcPackage) -> Result<Conformance> {
    Ok(inspect_graph(package)?.conformance)
}

/// Validate the presentation and table-style attachment graph without
/// allocating an owned copy of the catalog XML.
///
/// This is intended for package writers that need to preserve the optional
/// attachment while rebuilding the presentation part.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn present(package: &OpcPackage) -> Result<bool> {
    Ok(link(package)?.is_some())
}

/// Load the presentation's optional validated table-style catalog.
///
/// The returned list copies the bounded part payload exactly once so it can
/// outlive the package borrow and later move unchanged bytes through [`put`].
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage) -> Result<Option<List>> {
    let graph = inspect_graph(package)?;
    let Some(attachment) = graph.attachment else {
        return Ok(None);
    };
    let xml = own_blob(
        package.get_part(&attachment.part)?.blob(),
        "table-style XML ownership",
    )?;
    Ok(Some(parse_owned(xml)?))
}

/// Create or replace the presentation's table-style catalog atomically.
///
/// Loaded, unedited catalogs retain exact producer bytes. A byte-identical
/// load→put is a signature-preserving no-op; changed XML moves into a staged
/// part only after conformance, topology, and relationship checks succeed.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn put(package: &mut OpcPackage, list: List) -> Result<bool> {
    let graph = inspect_graph(package)?;
    if list.conformance != graph.conformance {
        return Err(invalid(
            "table-style conformance differs from the presentation",
        ));
    }
    let xml = list.into_xml()?;
    if let Some(attachment) = graph.attachment {
        let stored = package.get_part(&attachment.part)?.blob();
        if stored == xml || semantic_xml_eq(stored, &xml)? {
            return Ok(false);
        }
        let staged = BlobPart::new(attachment.part, ct::PML_TABLE_STYLES.into(), xml);
        package.add_part(Box::new(staged));
        package.unsign();
        return Ok(true);
    }

    let part_name = available_part_name(package)?;
    let relationship_id = {
        let presentation = package.get_part(&graph.presentation)?;
        available_relationship_id(presentation)?
    };
    let target = part_name.relative_ref(graph.presentation.base_uri());
    let mut staged_presentation = package.get_part(&graph.presentation)?.clone_part();
    staged_presentation.rels_mut().try_add_relationship(
        graph.conformance.relationship().into(),
        target,
        relationship_id,
        TargetMode::Internal,
    )?;
    package.validate_new_part_name(&part_name)?;
    let staged = BlobPart::new(part_name, ct::PML_TABLE_STYLES.into(), xml);

    package.add_part(Box::new(staged));
    package.add_part(staged_presentation);
    package.unsign();
    Ok(true)
}

/// Remove and return the optional table-style catalog atomically.
///
/// Absence is an idempotent, signature-preserving `Ok(None)`. A catalog with
/// any unexpected inbound edge is rejected before either relationship or part
/// is changed.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn remove(package: &mut OpcPackage) -> Result<Option<List>> {
    let graph = inspect_graph(package)?;
    let Some(attachment) = graph.attachment else {
        return Ok(None);
    };
    let xml = own_blob(
        package.get_part(&attachment.part)?.blob(),
        "removed table-style XML",
    )?;
    let list = parse_owned(xml)?;
    let mut staged_presentation = package.get_part(&graph.presentation)?.clone_part();
    if staged_presentation
        .rels_mut()
        .remove(&attachment.relationship_id)
        .is_none()
    {
        return Err(invalid(
            "validated table-style relationship disappeared before commit",
        ));
    }

    package.add_part(staged_presentation);
    let _ = package.remove_part(&attachment.part);
    package.unsign();
    Ok(Some(list))
}

struct Graph {
    presentation: PackURI,
    conformance: Conformance,
    attachment: Option<Attachment>,
}

struct Attachment {
    relationship_id: String,
    part: PackURI,
}

fn inspect_graph(package: &OpcPackage) -> Result<Graph> {
    if package.part_count() > MAX_GRAPH_PARTS {
        return Err(limit("table-style package parts", MAX_GRAPH_PARTS));
    }
    let mut main_relationship = None;
    for relationship in package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    }) {
        if main_relationship.replace(relationship).is_some() {
            return Err(invalid("package has multiple main-document relationships"));
        }
    }
    let main_relationship = main_relationship
        .ok_or_else(|| invalid("package main-document relationship is missing"))?;
    if main_relationship.is_external() {
        return Err(invalid(
            "package main-document relationship cannot be external",
        ));
    }
    let requested_presentation = main_relationship.target_partname()?;
    let presentation = package.get_part(&requested_presentation)?;
    require_presentation_content_type(presentation.content_type())?;
    let presentation_name = presentation.partname().clone();
    let conformance = presentation_conformance(presentation.blob())?;
    if main_relationship.reltype() != conformance.office_document() {
        return Err(invalid(
            "package main-document relationship conformance differs from the presentation",
        ));
    }
    let mut selected = None;
    for relationship in presentation
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::TABLE_STYLES | STRICT_REL))
    {
        if selected.is_some() {
            return Err(invalid(
                "presentation has multiple table-style relationships",
            ));
        }
        if relationship.reltype() != conformance.relationship() {
            return Err(invalid(
                "table-style relationship conformance differs from the presentation",
            ));
        }
        if relationship.is_external() {
            return Err(invalid("table-style relationship cannot be external"));
        }
        let requested = relationship.target_partname()?;
        let part = package.get_part(&requested)?;
        let part_name = part.partname().clone();
        validate_part_name(&part_name)?;
        if part.content_type() != ct::PML_TABLE_STYLES {
            return Err(Error::ContentType {
                expected: ct::PML_TABLE_STYLES.into(),
                actual: part.content_type().into(),
            });
        }
        if part.rels().iter().next().is_some() {
            return Err(invalid(
                "table-style part must not own package relationships",
            ));
        }
        let parsed = scan(part.blob())?;
        if parsed.conformance != conformance {
            return Err(invalid(
                "table-style XML conformance differs from the presentation",
            ));
        }
        selected = Some(Attachment {
            relationship_id: relationship.r_id().to_owned(),
            part: part_name,
        });
    }
    if let Some(attachment) = &selected {
        if package.iter_parts().any(|part| {
            part.content_type() == ct::PML_TABLE_STYLES && part.partname() != &attachment.part
        }) {
            return Err(invalid(
                "orphan table-style part exists outside the presentation relationship",
            ));
        }
        validate_inbound(package, &presentation_name, attachment)?;
    } else if package
        .iter_parts()
        .any(|part| part.content_type() == ct::PML_TABLE_STYLES)
    {
        return Err(invalid(
            "orphan table-style part exists without a presentation relationship",
        ));
    }
    Ok(Graph {
        presentation: presentation_name,
        conformance,
        attachment: selected,
    })
}

fn validate_inbound(
    package: &OpcPackage,
    presentation: &PackURI,
    attachment: &Attachment,
) -> Result<()> {
    let mut links = 0usize;
    let mut expected = 0usize;
    for relationship in package.rels().iter() {
        inspect_inbound(
            package,
            None,
            relationship,
            presentation,
            attachment,
            &mut expected,
        )?;
        links = checked_link_count(links)?;
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            inspect_inbound(
                package,
                Some(source.partname()),
                relationship,
                presentation,
                attachment,
                &mut expected,
            )?;
            links = checked_link_count(links)?;
        }
    }
    if expected != 1 {
        return Err(invalid(
            "table-style part must have exactly one presentation owner",
        ));
    }
    Ok(())
}

fn inspect_inbound(
    package: &OpcPackage,
    source: Option<&PackURI>,
    relationship: &litchi_opc::Relationship,
    presentation: &PackURI,
    attachment: &Attachment,
    expected: &mut usize,
) -> Result<()> {
    if relationship.is_external() {
        return Ok(());
    }
    let requested = relationship.target_partname()?;
    let targets_style = requested == attachment.part
        || package
            .get_part(&requested)
            .is_ok_and(|part| part.partname() == &attachment.part);
    if !targets_style {
        return Ok(());
    }
    if source == Some(presentation) && relationship.r_id() == attachment.relationship_id {
        *expected = expected
            .checked_add(1)
            .ok_or_else(|| invalid("table-style inbound count overflow"))?;
        return Ok(());
    }
    Err(invalid(format!(
        "table-style part '{}' has an unexpected inbound relationship '{}' from '{}'",
        attachment.part.as_str(),
        relationship.r_id(),
        source.map_or("package root", PackURI::as_str),
    )))
}

fn checked_link_count(value: usize) -> Result<usize> {
    let value = value
        .checked_add(1)
        .ok_or_else(|| limit("table-style package relationships", MAX_GRAPH_RELATIONSHIPS))?;
    if value > MAX_GRAPH_RELATIONSHIPS {
        Err(limit(
            "table-style package relationships",
            MAX_GRAPH_RELATIONSHIPS,
        ))
    } else {
        Ok(value)
    }
}

fn require_presentation_content_type(value: &str) -> Result<()> {
    if matches!(
        value,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: "a PresentationML presentation, slideshow, or template main part".into(),
            actual: value.into(),
        })
    }
}

fn presentation_conformance(xml: &[u8]) -> Result<Conformance> {
    if xml.len() > MAX_PRESENTATION_BYTES {
        return Err(limit("presentation XML bytes", MAX_PRESENTATION_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut profile = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut closed = false;
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                bump_presentation_node(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("presentation XML depth", MAX_DEPTH))?;
                if depth > MAX_DEPTH {
                    return Err(limit("presentation XML depth", MAX_DEPTH));
                }
                if depth == 1 {
                    if profile.is_some() || element.local_name().as_ref() != b"presentation" {
                        return Err(invalid("invalid presentation root"));
                    }
                    profile = presentation_namespace(&namespace);
                    if profile.is_none() {
                        return Err(invalid("presentation root has the wrong namespace"));
                    }
                }
            },
            Event::Empty(element) => {
                bump_presentation_node(&mut nodes)?;
                if depth == 0 {
                    if profile.is_some() || element.local_name().as_ref() != b"presentation" {
                        return Err(invalid("invalid presentation root"));
                    }
                    profile = presentation_namespace(&namespace);
                    if profile.is_none() {
                        return Err(invalid("presentation root has the wrong namespace"));
                    }
                    closed = true;
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("presentation XML nesting underflow"));
                }
                if depth == 1 {
                    let selected = profile.ok_or_else(|| invalid("missing presentation root"))?;
                    if element.local_name().as_ref() != b"presentation"
                        || presentation_namespace(&namespace) != Some(selected)
                    {
                        return Err(invalid("presentation root closes with a wrong element"));
                    }
                    closed = true;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "presentation XML must not contain a DTD or processing instruction",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !closed {
        return Err(invalid("presentation XML root is unterminated"));
    }
    profile.ok_or_else(|| invalid("missing presentation root"))
}

fn presentation_namespace(namespace: &ResolveResult<'_>) -> Option<Conformance> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == P.as_bytes() => {
            Some(Conformance::Transitional)
        },
        ResolveResult::Bound(Namespace(value)) if *value == PS.as_bytes() => {
            Some(Conformance::Strict)
        },
        _ => None,
    }
}

fn bump_presentation_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("presentation XML nodes", MAX_NODES))?;
    if *nodes > MAX_NODES {
        Err(limit("presentation XML nodes", MAX_NODES))
    } else {
        Ok(())
    }
}

fn validate_part_name(part: &PackURI) -> Result<()> {
    if part.as_str().starts_with("/ppt/")
        && part
            .as_str()
            .rsplit_once('.')
            .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("xml"))
    {
        Ok(())
    } else {
        Err(invalid("table-style part must be an XML part below /ppt/"))
    }
}

fn available_part_name(package: &OpcPackage) -> Result<PackURI> {
    for number in 1..=PART_NAME_ATTEMPTS {
        let path = if number == 1 {
            "/ppt/tableStyles.xml".to_owned()
        } else {
            format!("/ppt/tableStyles{number}.xml")
        };
        let candidate = PackURI::new(&path).map_err(Error::Invalid)?;
        if package.get_part(&candidate).is_err() {
            package.validate_new_part_name(&candidate)?;
            return Ok(candidate);
        }
    }
    Err(limit(
        "table-style part-name allocation attempts",
        PART_NAME_ATTEMPTS,
    ))
}

fn available_relationship_id(owner: &dyn Part) -> Result<String> {
    for number in 1..=RELATIONSHIP_ID_ATTEMPTS {
        let candidate = format!("rId{number}");
        if owner.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(limit(
        "table-style relationship-ID allocation attempts",
        RELATIONSHIP_ID_ATTEMPTS,
    ))
}

fn own_blob(blob: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(blob.len())
        .map_err(|source| allocation(resource, source))?;
    owned.extend_from_slice(blob);
    Ok(owned)
}
pub(super) fn validate_parsed(parsed: &Parsed) -> Result<()> {
    if parsed.defs.len() > MAX_STYLES {
        return Err(limit("table-style count", MAX_STYLES));
    }
    let mut ids = HashSet::new();
    ids.try_reserve(parsed.defs.len())
        .map_err(|source| allocation("table-style ID validation", source))?;
    for style in &parsed.defs {
        validate_name(&style.name)?;
        if !ids.insert(style.id) {
            return Err(invalid(format!("duplicate table-style ID {}", style.id)));
        }
    }
    Ok(())
}

pub(super) fn validate_list(list: &List) -> Result<()> {
    if list.defs.len() > MAX_STYLES {
        return Err(limit("table-style count", MAX_STYLES));
    }
    let mut ids = HashSet::new();
    ids.try_reserve(list.defs.len())
        .map_err(|source| allocation("table-style ID validation", source))?;
    for style in &list.defs {
        validate_def(style)?;
        if !ids.insert(style.id) {
            return Err(invalid(format!("duplicate table-style ID {}", style.id)));
        }
    }
    Ok(())
}

pub(super) fn validate_def(style: &Def) -> Result<()> {
    validate_name(&style.name)?;
    if style.parts.bits() & !Parts::all().bits() != 0 {
        return Err(invalid("table style contains unknown region flags"));
    }
    Ok(())
}

pub(super) fn validate_name(name: &str) -> Result<()> {
    if name.len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit("table-style name bytes", MAX_ATTRIBUTE_BYTES));
    }
    if !name.chars().all(xml_char) {
        return Err(invalid(
            "table-style name contains an invalid XML character",
        ));
    }
    Ok(())
}
