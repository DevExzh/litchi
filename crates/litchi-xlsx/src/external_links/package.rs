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
use super::model::{
    AlternateUrl, AlternateUrls, Conformance, Link, MAX_CACHE_TEXT_BYTES,
    MAX_EXTERNAL_TARGET_BYTES, Target,
};
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
    if let Link::Workbook(link) = kind {
        add_alternate_target_relationships(&mut part, link.alternate_urls.as_ref())?;
    }
    Ok(part)
}

fn add_alternate_target_relationships(
    part: &mut BlobPart,
    alternate_urls: Option<&AlternateUrls>,
) -> Result<()> {
    let Some(alternate_urls) = alternate_urls else {
        return Ok(());
    };
    let mut seen = HashMap::<String, (String, String)>::new();
    for url in [
        alternate_urls.absolute_url.as_ref(),
        alternate_urls.relative_url.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        let metadata = (url.target.clone(), url.relationship_type.clone());
        if let Some(previous) = seen.insert(url.relationship_id.clone(), metadata.clone()) {
            if previous != metadata {
                return Err(invalid(format!(
                    "alternate URL relationship ID '{}' has conflicting targets",
                    url.relationship_id
                )));
            }
            continue;
        }
        if part.rels().get(&url.relationship_id).is_some() {
            continue;
        }
        part.rels_mut().add_relationship(
            url.relationship_type.clone(),
            url.target.clone(),
            url.relationship_id.clone(),
            true,
        );
    }
    Ok(())
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
    if part.blob().len() > MAX_CACHE_TEXT_BYTES {
        return Err(limit("external-link XML"));
    }
    let mut kind = parse_external_link(part.blob())?;
    match &mut kind {
        Link::Workbook(book) => {
            resolve_external_target(part, &mut book.target, "externalBook")?;
            if let Some(alternate_urls) = &mut book.alternate_urls {
                if let Some(url) = &mut alternate_urls.absolute_url {
                    resolve_external_target(part, url, "absoluteUrl")?;
                }
                if let Some(url) = &mut alternate_urls.relative_url {
                    resolve_external_target(part, url, "relativeUrl")?;
                }
            }
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

trait ResolvableTarget {
    fn relationship_id(&self) -> &str;
    fn set_target(&mut self, target: String, relationship_type: String);
}

impl ResolvableTarget for Target {
    fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    fn set_target(&mut self, target: String, relationship_type: String) {
        self.target = target;
        self.relationship_type = relationship_type;
    }
}

impl ResolvableTarget for AlternateUrl {
    fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    fn set_target(&mut self, target: String, relationship_type: String) {
        self.target = target;
        self.relationship_type = relationship_type;
    }
}

fn resolve_external_target(
    part: &dyn Part,
    target: &mut impl ResolvableTarget,
    description: &str,
) -> Result<()> {
    let relationship = part.rels().get(target.relationship_id()).ok_or_else(|| {
        invalid(format!(
            "{description} references missing relationship '{}'",
            target.relationship_id()
        ))
    })?;
    if !relationship.is_external() {
        return Err(invalid(format!(
            "{description} target relationship must be external"
        )));
    }
    if !is_external_workbook_relationship(relationship.reltype()) {
        return Err(invalid(format!(
            "{description} target has invalid relationship type '{}'",
            relationship.reltype()
        )));
    }
    let target_ref = relationship.target_ref();
    if target_ref.len() > MAX_EXTERNAL_TARGET_BYTES {
        return Err(limit(&format!("{description} target URI")));
    }
    if target_ref.chars().any(|character| {
        character.is_control() || character == '\u{fffe}' || character == '\u{ffff}'
    }) {
        return Err(invalid(format!(
            "{description} target URI contains an invalid character"
        )));
    }
    target.set_target(target_ref.to_owned(), relationship.reltype().to_string());
    Ok(())
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
    // Direct package APIs share this path with transactions.  Stage every
    // mutation on a clone so a late graph/source failure cannot leave a
    // caller-visible package with changed XML or relationships.
    let mut candidate = package.clone();
    apply_entries_in_place(&mut candidate, before, after, conformance)?;
    *package = candidate;
    Ok(())
}

fn apply_entries_in_place(
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

        match package.get_part(&entry.part_uri) {
            Ok(_) => {
                return Err(invalid(format!(
                    "new external-link part '{}' already exists",
                    entry.part_uri
                )));
            },
            Err(litchi_opc::OpcError::PartNotFound(_)) => {},
            Err(error) => return Err(error.into()),
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
    if part.blob().len() > MAX_CACHE_TEXT_BYTES {
        return Err(limit("external-link XML"));
    }
    let source = part.blob().to_vec();
    let updated = patch_source(&source, &before.link, &after.link, conformance)?;
    let current_targets = target_map(&before.link)?;
    let next_targets = target_map(&after.link)?;

    let part = package.get_part_mut(&before.part_uri)?;
    validate_target_relationship_update(part, &current_targets, &next_targets)?;
    part.set_blob(updated);
    update_target_relationships(part, &current_targets, &next_targets)?;
    Ok(())
}

fn validate_target_relationship_update(
    part: &dyn Part,
    before: &HashMap<String, Target>,
    after: &HashMap<String, Target>,
) -> Result<()> {
    for id in after.keys() {
        if part.rels().get(id).is_some() {
            let represented_before = before.contains_key(id);
            if !represented_before {
                return Err(invalid(format!(
                    "external-link target relationship ID '{}' is already in use",
                    id
                )));
            }
        }
    }
    for id in before.keys() {
        if after.contains_key(id) {
            continue;
        }
        if part.rels().get(id).is_none() {
            return Err(invalid(format!(
                "external-link target relationship ID '{}' is missing",
                id
            )));
        }
    }
    Ok(())
}

fn update_target_relationships(
    part: &mut dyn Part,
    before: &HashMap<String, Target>,
    after: &HashMap<String, Target>,
) -> Result<()> {
    for id in before.keys().filter(|id| !after.contains_key(*id)) {
        part.rels_mut().remove(id);
    }
    for (id, target) in after {
        let unchanged = before.get(id).is_some_and(|old| old == target);
        if unchanged {
            continue;
        }
        // `Relationships::add_relationship` intentionally preserves an
        // existing ID for compatibility.  This owner is replacing an owned
        // relationship, so remove the old physical record first; otherwise a
        // successful semantic edit would continue resolving the stale target.
        if before.contains_key(id) {
            part.rels_mut().remove(id);
        }
        part.rels_mut().add_relationship(
            target.relationship_type.clone(),
            target.target.clone(),
            id.clone(),
            true,
        );
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

fn target_map(link: &Link) -> Result<HashMap<String, Target>> {
    let mut targets = HashMap::new();
    let values = match link {
        Link::Workbook(link) => {
            let mut values = vec![link.target.clone()];
            if let Some(alternate_urls) = &link.alternate_urls {
                if let Some(url) = &alternate_urls.absolute_url {
                    values.push(Target {
                        relationship_id: url.relationship_id.clone(),
                        target: url.target.clone(),
                        relationship_type: url.relationship_type.clone(),
                    });
                }
                if let Some(url) = &alternate_urls.relative_url {
                    values.push(Target {
                        relationship_id: url.relationship_id.clone(),
                        target: url.target.clone(),
                        relationship_type: url.relationship_type.clone(),
                    });
                }
            }
            values
        },
        Link::Dde(_) => Vec::new(),
        Link::Ole(link) => vec![link.target.clone()],
    };
    for target in values {
        if let Some(previous) = targets.insert(target.relationship_id.clone(), target.clone())
            && previous != target
        {
            return Err(invalid(format!(
                "external-link target relationship ID '{}' has conflicting targets",
                target.relationship_id
            )));
        }
    }
    Ok(targets)
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
