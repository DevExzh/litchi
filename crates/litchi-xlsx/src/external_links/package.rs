//! OPC relationship and part orchestration for external links.

use std::collections::{HashMap, HashSet};

use crate::error::Result;
use litchi_ooxml_common::external_link::{
    EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES, is_external_workbook_relationship,
};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

use super::codec::{parse_external_link, patch_source};
use super::model::{Conformance, Link, MAX_EXTERNAL_TARGET_BYTES, Target};
use super::{invalid, limit, validation};

/// One external-link part together with its workbook package relationship.
///
/// This physical identity belongs to the OPC package layer; the typed link
/// models remain usable without a package or relationship catalog.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    pub index: u32,
    pub relationship_id: String,
    pub part_uri: PackURI,
    pub link: Link,
}

pub fn build_external_link_part(part_uri: PackURI, kind: &Link) -> Result<BlobPart> {
    build_external_link_part_with_conformance(part_uri, kind, Conformance::Transitional)
}

pub fn build_external_link_part_with_conformance(
    part_uri: PackURI,
    kind: &Link,
    conformance: Conformance,
) -> Result<BlobPart> {
    let xml = kind.to_xml_with_conformance(conformance)?;
    let mut part = BlobPart::new(
        part_uri,
        litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
        xml,
    );
    match kind {
        Link::Workbook(link) => add_external_target_relationship(
            &mut part,
            &link.target,
            EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES,
            "external workbook",
        )?,
        Link::Ole(link) => add_external_target_relationship(
            &mut part,
            &link.target,
            &[rt::OLE_OBJECT, rt::STRICT_OLE_OBJECT],
            "OLE",
        )?,
        Link::Dde(_) => {},
    }
    Ok(part)
}

trait TargetMetadata {
    fn relationship_id(&self) -> &str;
    fn target(&self) -> &str;
    fn relationship_type(&self) -> &str;
}

impl TargetMetadata for Target {
    fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
    fn target(&self) -> &str {
        &self.target
    }
    fn relationship_type(&self) -> &str {
        &self.relationship_type
    }
}

fn add_external_target_relationship(
    part: &mut BlobPart,
    target: &impl TargetMetadata,
    allowed_types: &[&str],
    description: &str,
) -> Result<()> {
    validate_external_target(target, allowed_types, description)?;
    part.rels_mut().add_relationship(
        target.relationship_type().to_string(),
        target.target().to_string(),
        target.relationship_id().to_string(),
        true,
    );
    Ok(())
}

fn validate_external_target(
    target: &impl TargetMetadata,
    allowed_types: &[&str],
    description: &str,
) -> Result<()> {
    if target.relationship_id().is_empty() {
        return Err(invalid(format!(
            "{description} relationship ID must not be empty"
        )));
    }
    if target.target().is_empty() {
        return Err(invalid(format!("{description} target must not be empty")));
    }
    if target.target().len() > MAX_EXTERNAL_TARGET_BYTES {
        return Err(limit(&format!("{description} target URI")));
    }
    if target.target().chars().any(|character| {
        character.is_control() || character == '\u{fffe}' || character == '\u{ffff}'
    }) {
        return Err(invalid(format!(
            "{description} target URI contains an invalid character"
        )));
    }
    if target.relationship_id().len() > 1024
        || target.relationship_id().chars().any(char::is_control)
    {
        return Err(invalid(format!("{description} relationship ID is invalid")));
    }
    if !allowed_types.contains(&target.relationship_type()) {
        return Err(invalid(format!(
            "{description} has invalid relationship type '{}'",
            target.relationship_type()
        )));
    }
    Ok(())
}

pub fn load_external_link(
    part: &dyn Part,
    workbook_relationship_id: String,
    index: u32,
) -> Result<Entry> {
    let mut kind = parse_external_link(part.blob())?;
    match &mut kind {
        Link::Workbook(book) => {
            let relationship = part
                .rels()
                .get(&book.target.relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "externalBook references missing relationship '{}'",
                        book.target.relationship_id
                    ))
                })?;
            if !relationship.is_external() {
                return Err(invalid("externalBook target relationship must be external"));
            }
            if !is_external_workbook_relationship(relationship.reltype()) {
                return Err(invalid(format!(
                    "externalBook target has invalid relationship type '{}'",
                    relationship.reltype()
                )));
            }
            book.target.target = relationship.target_ref().to_string();
            book.target.relationship_type = relationship.reltype().to_string();
        },
        Link::Dde(_) => {},
        Link::Ole(link) => {
            let relationship = part
                .rels()
                .get(&link.target.relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "oleLink references missing relationship '{}'",
                        link.target.relationship_id
                    ))
                })?;
            if !relationship.is_external() {
                return Err(invalid("oleLink target relationship must be external"));
            }
            if !matches!(
                relationship.reltype(),
                rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT
            ) {
                return Err(invalid(format!(
                    "oleLink target has invalid relationship type '{}'",
                    relationship.reltype()
                )));
            }
            link.target.target = relationship.target_ref().to_string();
            link.target.relationship_type = relationship.reltype().to_string();
        },
    }
    Ok(Entry {
        index,
        relationship_id: workbook_relationship_id,
        part_uri: part.partname().clone(),
        link: kind,
    })
}

/// Load every external-link part owned by the workbook.
pub fn load_external_links(package: &OpcPackage) -> Result<Vec<Entry>> {
    validate_graph(package)?;
    let workbook = package.main_document_part()?;
    require_workbook(workbook.content_type())?;

    let mut relationships = workbook
        .rels()
        .iter()
        .filter(|relationship| is_owner_relationship(relationship.reltype()))
        .collect::<Vec<_>>();
    relationships.sort_by(|left, right| left.r_id().cmp(right.r_id()));

    let mut entries = Vec::with_capacity(relationships.len());
    for (index, relationship) in relationships.into_iter().enumerate() {
        if relationship.is_external() {
            return Err(invalid(
                "workbook external-link relationship must target an internal part",
            ));
        }
        let part_uri = relationship.target_partname()?;
        let part = package.get_part(&part_uri)?;
        if part.content_type() != litchi_opc::constants::content_type::SML_EXTERNAL_LINK {
            return Err(invalid(format!(
                "external-link relationship '{}' targets part '{}' with content type '{}'",
                relationship.r_id(),
                part_uri,
                part.content_type()
            )));
        }
        let index =
            u32::try_from(index).map_err(|_source| invalid("external-link index exceeds u32"))?;
        entries.push(load_external_link(
            part,
            relationship.r_id().to_owned(),
            index,
        )?);
    }
    Ok(entries)
}

/// Store an initial external-link catalog on a workbook with no existing
/// external-link owner. Every target remains an inert OPC relationship.
pub fn store_external_links(
    package: &mut OpcPackage,
    links: &[Link],
    conformance: Conformance,
) -> Result<Vec<Entry>> {
    let before = load_external_links(package)?;
    if !before.is_empty() {
        return Err(invalid(
            "workbook already contains external links; use a transaction to edit the catalog",
        ));
    }
    let entries = allocate_entries(package, links)?;
    validation::entries(&entries, conformance)?;
    apply_entries(package, &[], &entries, conformance)?;
    load_external_links(package)
}

/// Add one external link to the workbook and return its physical entry.
pub fn add_external_link(
    package: &mut OpcPackage,
    link: Link,
    conformance: Conformance,
) -> Result<Entry> {
    let before = load_external_links(package)?;
    let mut after = before.clone();
    let entry = allocate_entries(package, std::slice::from_ref(&link))?
        .into_iter()
        .next()
        .ok_or_else(|| invalid("failed to allocate external-link entry"))?;
    after.push(entry.clone());
    validation::entries(&after, conformance)?;
    apply_entries(package, &before, &after, conformance)?;
    load_external_links(package)?
        .into_iter()
        .find(|candidate| candidate.relationship_id == entry.relationship_id)
        .ok_or_else(|| invalid("published external-link entry is absent"))
}

/// Replace one external link while retaining its workbook relationship and
/// part identity.
pub fn replace_external_link(
    package: &mut OpcPackage,
    index: usize,
    link: Link,
    conformance: Conformance,
) -> Result<Entry> {
    let before = load_external_links(package)?;
    let mut after = before.clone();
    let entry = after
        .get_mut(index)
        .ok_or_else(|| invalid(format!("external-link index {index} is absent")))?;
    entry.link = link;
    validation::entries(&after, conformance)?;
    apply_entries(package, &before, &after, conformance)?;
    load_external_links(package)?
        .into_iter()
        .nth(index)
        .ok_or_else(|| invalid("published external-link entry is absent"))
}

/// Remove one external link and its unreferenced owned part.
pub fn remove_external_link(package: &mut OpcPackage, index: usize) -> Result<Option<Entry>> {
    let before = load_external_links(package)?;
    if index >= before.len() {
        return Ok(None);
    }
    let removed = before[index].clone();
    let mut after = before.clone();
    after.remove(index);
    apply_entries(package, &before, &after, Conformance::Transitional)?;
    Ok(Some(removed))
}

/// Validate workbook ownership, nested target relationships, and orphan link
/// parts without opening or refreshing any external target.
pub fn validate_graph(package: &OpcPackage) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_owner_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "package root cannot source a workbook external-link relationship",
        ));
    }

    let workbook = package.main_document_part()?;
    require_workbook(workbook.content_type())?;
    let mut target_parts = HashSet::new();
    let mut owner_ids = HashSet::new();
    for relationship in workbook
        .rels()
        .iter()
        .filter(|relationship| is_owner_relationship(relationship.reltype()))
    {
        if !owner_ids.insert(relationship.r_id().to_owned()) {
            return Err(invalid(format!(
                "duplicate workbook external-link relationship '{}'",
                relationship.r_id()
            )));
        }
        if relationship.is_external() {
            return Err(invalid(
                "workbook external-link relationship cannot be external",
            ));
        }
        let target = relationship.target_partname()?;
        if !target_parts.insert(target.to_string()) {
            return Err(invalid(format!(
                "external-link part '{target}' is targeted more than once"
            )));
        }
        let part = package.get_part(&target)?;
        if part.content_type() != litchi_opc::constants::content_type::SML_EXTERNAL_LINK {
            return Err(invalid(format!(
                "external-link target '{}' has invalid content type '{}'",
                target,
                part.content_type()
            )));
        }
        load_external_link(part, relationship.r_id().to_owned(), 0)?;
    }

    for part in package.iter_parts().filter(|part| {
        part.content_type() == litchi_opc::constants::content_type::SML_EXTERNAL_LINK
    }) {
        if !target_parts.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "external-link part '{}' has no workbook owner",
                part.partname()
            )));
        }
    }
    Ok(())
}

pub(crate) fn apply_entries(
    package: &mut OpcPackage,
    before: &[Entry],
    after: &[Entry],
    conformance: Conformance,
) -> Result<()> {
    validation::entries(after, conformance)?;
    let workbook_uri = package.main_document_part()?.partname().clone();
    let workbook = package.get_part(&workbook_uri)?;
    require_workbook(workbook.content_type())?;

    let before_by_identity = before
        .iter()
        .map(|entry| {
            (
                (entry.relationship_id.clone(), entry.part_uri.clone()),
                entry,
            )
        })
        .collect::<HashMap<_, _>>();
    let after_identities = after
        .iter()
        .map(|entry| (entry.relationship_id.clone(), entry.part_uri.clone()))
        .collect::<HashSet<_>>();

    // Remove deleted owners first. Their parts are removed only when no
    // package relationship still references them.
    for entry in before {
        if after_identities.contains(&(entry.relationship_id.clone(), entry.part_uri.clone())) {
            continue;
        }
        let workbook = package.get_part_mut(&workbook_uri)?;
        workbook.rels_mut().remove(&entry.relationship_id);
        if !part_is_referenced(package, &entry.part_uri) {
            package.remove_part(&entry.part_uri);
        }
    }

    for entry in after {
        let identity = (entry.relationship_id.clone(), entry.part_uri.clone());
        if let Some(previous) = before_by_identity.get(&identity) {
            if previous.link != entry.link {
                replace_existing_part(package, previous, entry, conformance)?;
            }
            continue;
        }

        if package.get_part(&entry.part_uri).is_ok() {
            return Err(invalid(format!(
                "new external-link part '{}' already exists",
                entry.part_uri
            )));
        }
        let part = build_external_link_part_with_conformance(
            entry.part_uri.clone(),
            &entry.link,
            conformance,
        )?;
        package.try_add_part(Box::new(part))?;
        if package
            .get_part(&workbook_uri)?
            .rels()
            .get(&entry.relationship_id)
            .is_some()
        {
            return Err(invalid(format!(
                "workbook relationship ID '{}' already exists",
                entry.relationship_id
            )));
        }
        package
            .get_part_mut(&workbook_uri)?
            .rels_mut()
            .add_relationship(
                conformance.external_link_relationship().to_owned(),
                entry.part_uri.relative_ref(workbook_uri.base_uri()),
                entry.relationship_id.clone(),
                false,
            );
    }

    validate_graph(package)
}

fn replace_existing_part(
    package: &mut OpcPackage,
    before: &Entry,
    after: &Entry,
    conformance: Conformance,
) -> Result<()> {
    validation::link(&after.link, conformance)?;
    let part = package.get_part(&before.part_uri)?;
    if part.content_type() != litchi_opc::constants::content_type::SML_EXTERNAL_LINK {
        return Err(invalid("external-link replacement targets a non-link part"));
    }
    let source = part.blob().to_vec();
    let updated = patch_source(&source, &before.link, &after.link, conformance)?;
    let mut current_target = target_of(&before.link).cloned();
    let next_target = target_of(&after.link).cloned();

    let part = package.get_part_mut(&before.part_uri)?;
    part.set_blob(updated);
    update_target_relationship(part, &mut current_target, next_target.as_ref())?;
    Ok(())
}

fn update_target_relationship(
    part: &mut dyn Part,
    before: &mut Option<Target>,
    after: Option<&Target>,
) -> Result<()> {
    match (before.take(), after) {
        (Some(old), Some(new)) if old.relationship_id == new.relationship_id => {
            part.rels_mut().add_relationship(
                new.relationship_type.clone(),
                new.target.clone(),
                new.relationship_id.clone(),
                true,
            );
        },
        (Some(old), Some(new)) => {
            if part.rels().get(&new.relationship_id).is_some()
                && new.relationship_id != old.relationship_id
            {
                return Err(invalid(format!(
                    "external-link target relationship ID '{}' is already in use",
                    new.relationship_id
                )));
            }
            part.rels_mut().remove(&old.relationship_id);
            part.rels_mut().add_relationship(
                new.relationship_type.clone(),
                new.target.clone(),
                new.relationship_id.clone(),
                true,
            );
        },
        (Some(old), None) => {
            part.rels_mut().remove(&old.relationship_id);
        },
        (None, Some(new)) => {
            if part.rels().get(&new.relationship_id).is_some() {
                return Err(invalid(format!(
                    "external-link target relationship ID '{}' is already in use",
                    new.relationship_id
                )));
            }
            part.rels_mut().add_relationship(
                new.relationship_type.clone(),
                new.target.clone(),
                new.relationship_id.clone(),
                true,
            );
        },
        (None, None) => {},
    }
    Ok(())
}

fn allocate_entries(package: &OpcPackage, links: &[Link]) -> Result<Vec<Entry>> {
    let workbook = package.main_document_part()?;
    let mut relationship_number = 1u32;
    let mut relationship_ids = workbook
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<HashSet<_>>();
    let mut part_number = 1u32;
    let mut part_names = package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect::<HashSet<_>>();
    let mut entries = Vec::with_capacity(links.len());
    for (index, link) in links.iter().cloned().enumerate() {
        let relationship_id = loop {
            let candidate = format!("rId{relationship_number}");
            relationship_number = relationship_number.saturating_add(1);
            if relationship_ids.insert(candidate.clone()) {
                break candidate;
            }
        };
        let part_uri = loop {
            let candidate = format!("/xl/externalLinks/externalLink{part_number}.xml");
            part_number = part_number.saturating_add(1);
            if part_names.insert(candidate.clone()) {
                break PackURI::new(candidate).map_err(invalid)?;
            }
        };
        let index =
            u32::try_from(index).map_err(|_source| invalid("external-link index exceeds u32"))?;
        entries.push(Entry {
            index,
            relationship_id,
            part_uri,
            link,
        });
    }
    Ok(entries)
}

fn target_of(link: &Link) -> Option<&Target> {
    match link {
        Link::Workbook(link) => Some(&link.target),
        Link::Dde(_) => None,
        Link::Ole(link) => Some(&link.target),
    }
}

fn part_is_referenced(package: &OpcPackage, part_uri: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
            .any(|relationship| {
                relationship
                    .target_partname()
                    .is_ok_and(|target| target == *part_uri)
            })
    })
}

fn require_workbook(content_type: &str) -> Result<()> {
    if content_type.starts_with("application/vnd.openxmlformats-officedocument.spreadsheetml.")
        || content_type.starts_with("application/vnd.ms-excel.")
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "part content type '{content_type}' is not an XLSX workbook"
        )))
    }
}

fn is_owner_relationship(value: &str) -> bool {
    matches!(value, rt::EXTERNAL_LINK | rt::STRICT_EXTERNAL_LINK)
}
