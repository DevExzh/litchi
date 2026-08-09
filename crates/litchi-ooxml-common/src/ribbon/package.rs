//! OPC package graph discovery and mutation for shared Ribbon customizations.

use crate::{Error, Result};
use litchi_opc::{OpcPackage, PackURI, Part, XmlPart};
use std::cmp::Ordering;
use std::collections::VecDeque;

use super::codec::{require_content_type, require_internal_target, validate_images, validate_xml};
use super::model::{Family, Limits, Set, Ui, Version};

pub(super) const CONTENT_TYPE: &str = "application/xml";
const PART_NAME_ATTEMPTS: usize = 10_000;
const MAX_GRAPH_LINKS: usize = 1_000_000;
const MAX_PART_NAMES: usize = 100_000;
const MAX_PART_NAME_BYTES: usize = 64 * 1024 * 1024;
const MAX_IMAGE_GC_EDGES: usize = 262_144;

/// Read both Ribbon family slots with safe default limits.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn load(package: &OpcPackage) -> Result<Set<'_>> {
    load_with(package, &Limits::standard())
}

/// Read both Ribbon family slots with explicit resource limits.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn load_with<'a>(package: &'a OpcPackage, limits: &Limits) -> Result<Set<'a>> {
    let mut scanned = 0usize;
    reject_part_sourced_ribbons(package, &mut scanned)?;

    let mut result = Set::default();
    let mut images = 0usize;
    for relationship in package.rels().iter() {
        bump_graph_links(&mut scanned)?;
        let Some(family) = Family::from_relationship(relationship.reltype()) else {
            continue;
        };
        result.require_empty(family)?;
        require_internal_target(relationship, "Ribbon")?;
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid Ribbon relationship target: {error}"))
        })?;
        let part = package.get_part(&target).map_err(|error| {
            Error::Missing(format!(
                "Ribbon part '{}' does not exist: {error}",
                target.as_str()
            ))
        })?;
        require_content_type(part, CONTENT_TYPE)?;
        let version = validate_xml(part.blob(), family, limits)?;
        validate_images(package, part, &mut images, limits)?;
        result.insert(Ui::new(
            part.partname(),
            relationship.r_id(),
            version,
            part.blob(),
        ));
    }
    Ok(result)
}

/// Create or replace one Ribbon family using safe default limits.
///
/// The XML allocation is moved into the OPC part. A byte-identical update is
/// a true no-op and preserves package signatures. Input must be UTF-8 XML.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn put(package: &mut OpcPackage, version: Version, xml: Vec<u8>) -> Result<()> {
    put_with(package, version, xml, &Limits::standard())
}

/// Create or replace one Ribbon family using explicit resource limits.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn put_with(
    package: &mut OpcPackage,
    version: Version,
    xml: Vec<u8>,
    limits: &Limits,
) -> Result<()> {
    let parsed = validate_xml(&xml, version.family(), limits)?;
    if parsed != version {
        return Err(Error::Invalid(format!(
            "Ribbon root namespace does not match requested {version:?} version"
        )));
    }

    let existing = {
        let ribbons = load_with(package, limits)?;
        match ribbons.get(version.family()) {
            Some(current) if current.version() == version && current.xml() == xml => return Ok(()),
            Some(current) => {
                let part = current.part().clone();
                let id = current.id().to_owned();
                let shared = has_other_inbound(package, &part, Some(&id))?;
                Some((part, id, shared))
            },
            None => None,
        }
    };

    if let Some((part, _, false)) = existing.as_ref() {
        package.get_part_mut(part)?.set_blob(xml);
        package.unsign();
        return Ok(());
    }

    let part = available_part_name(package, version)?;
    let target = part.as_str().trim_start_matches('/').to_owned();
    let mut replacement = XmlPart::new(part, CONTENT_TYPE.to_owned(), xml);
    if let Some((shared, _, true)) = existing.as_ref() {
        let base = replacement.partname().base_uri().to_owned();
        for relationship in package.get_part(shared)?.rels().iter() {
            let image = relationship.target_partname()?;
            replacement.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                image.relative_ref(&base),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )?;
        }
    }
    let relationship_id = existing.map(|(_, id, _)| id);
    if let Some(id) = relationship_id.as_deref()
        && package.rels_mut().remove(id).is_none()
    {
        return Err(Error::Relationship(format!(
            "Ribbon relationship '{id}' disappeared before commit"
        )));
    }
    package.add_part(Box::new(replacement));
    if let Some(id) = relationship_id {
        package
            .rels_mut()
            .add_relationship(version.relationship().to_owned(), target, id, false);
    } else {
        package.relate_to(&target, version.relationship());
    }
    package.unsign();
    Ok(())
}

/// Remove one Ribbon family and collect only parts that become unreferenced.
///
/// The complete deletion plan is resolved before mutation. Shared Ribbon or
/// image parts remain in the package. An absent family returns `Ok(false)` and
/// preserves signatures.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn remove(package: &mut OpcPackage, family: Family) -> Result<bool> {
    let (relationship_id, ribbon_part, images) = {
        let ribbons = load(package)?;
        let Some(selected) = ribbons.get(family) else {
            return Ok(false);
        };
        let ribbon_part = selected.part().clone();
        let relationship_id = selected.id().to_owned();
        let remove_part =
            !has_other_inbound(package, &ribbon_part, Some(relationship_id.as_str()))?;
        let images = if remove_part {
            let candidates = ribbon_images(package, &ribbon_part)?;
            removable_images(package, &ribbon_part, &candidates)?
        } else {
            Vec::new()
        };
        (relationship_id, remove_part.then_some(ribbon_part), images)
    };

    if package.rels_mut().remove(&relationship_id).is_none() {
        return Err(Error::Relationship(format!(
            "Ribbon relationship '{relationship_id}' disappeared before commit"
        )));
    }
    if let Some(part) = ribbon_part {
        let _ = package.remove_part(&part);
        for image in images {
            let _ = package.remove_part(&image);
        }
    }
    package.unsign();
    Ok(true)
}

fn reject_part_sourced_ribbons(package: &OpcPackage, scanned: &mut usize) -> Result<()> {
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            bump_graph_links(scanned)?;
            if Family::from_relationship(relationship.reltype()).is_some() {
                return Err(Error::Relationship(format!(
                    "Ribbon relationship '{}' must be sourced by the package, not part '{}'",
                    relationship.r_id(),
                    part.partname().as_str()
                )));
            }
        }
    }
    Ok(())
}

fn available_part_name(package: &OpcPackage, version: Version) -> Result<PackURI> {
    let existing = sorted_part_names(package)?;
    for suffix in 0..PART_NAME_ATTEMPTS {
        let path = if suffix == 0 {
            version.default_part().to_owned()
        } else {
            format!("/customUI/customUI{suffix}.xml")
        };
        let candidate = PackURI::new(&path)
            .map_err(|error| Error::Uri(format!("Ribbon part URI '{path}': {error}")))?;
        let folded = path.to_ascii_lowercase();
        if !part_name_conflicts(&existing, &folded) {
            package.validate_new_part_name(&candidate)?;
            return Ok(candidate);
        }
    }
    Err(Error::Invalid(
        "unable to allocate a unique Ribbon part name".into(),
    ))
}

fn sorted_part_names(package: &OpcPackage) -> Result<Vec<String>> {
    let mut names = Vec::with_capacity(package.part_count().min(MAX_PART_NAMES));
    let mut bytes = 0usize;
    for part in package.iter_parts() {
        let count = names.len().checked_add(1).ok_or(Error::Limit {
            resource: "Ribbon part-name allocation scan",
            max: MAX_PART_NAMES,
            actual: usize::MAX,
        })?;
        if count > MAX_PART_NAMES {
            return Err(Error::Limit {
                resource: "Ribbon part-name allocation scan",
                max: MAX_PART_NAMES,
                actual: count,
            });
        }
        bytes = bytes
            .checked_add(part.partname().as_str().len())
            .ok_or(Error::Limit {
                resource: "Ribbon part-name allocation bytes",
                max: MAX_PART_NAME_BYTES,
                actual: usize::MAX,
            })?;
        if bytes > MAX_PART_NAME_BYTES {
            return Err(Error::Limit {
                resource: "Ribbon part-name allocation bytes",
                max: MAX_PART_NAME_BYTES,
                actual: bytes,
            });
        }
        names.push(part.partname().as_str().to_ascii_lowercase());
    }
    names.sort_unstable();
    names.dedup();
    Ok(names)
}

fn part_name_conflicts(existing: &[String], candidate: &str) -> bool {
    if sorted_name_exists(existing, candidate) {
        return true;
    }
    for (index, _) in candidate.match_indices('/').skip(1) {
        if sorted_name_exists(existing, &candidate[..index]) {
            return true;
        }
    }
    let descendant_prefix = format!("{candidate}/");
    let position = existing.partition_point(|name| name.as_str() < descendant_prefix.as_str());
    existing
        .get(position)
        .is_some_and(|name| name.starts_with(&descendant_prefix))
}

fn sorted_name_exists(existing: &[String], wanted: &str) -> bool {
    existing
        .binary_search_by(|name| name.as_str().cmp(wanted))
        .is_ok()
}

fn ribbon_images(package: &OpcPackage, ribbon: &PackURI) -> Result<Vec<PackURI>> {
    let part = package.get_part(ribbon)?;
    let mut images = Vec::new();
    for relationship in part.rels().iter() {
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid Ribbon image target: {error}"))
        })?;
        let target = package.get_part(&target)?.partname();
        images.push(target.clone());
    }
    images.sort_unstable_by(compare_names);
    images.dedup_by(|left, right| same_name(left, right));
    Ok(images)
}

fn has_other_inbound(
    package: &OpcPackage,
    target: &PackURI,
    skipped_package_relationship: Option<&str>,
) -> Result<bool> {
    let mut scanned = 0usize;
    for relationship in package.rels().iter() {
        bump_graph_links(&mut scanned)?;
        if skipped_package_relationship == Some(relationship.r_id()) || relationship.is_external() {
            continue;
        }
        let related = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid package relationship target: {error}"))
        })?;
        if same_name(&related, target) {
            return Ok(true);
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            bump_graph_links(&mut scanned)?;
            if relationship.is_external() {
                continue;
            }
            let related = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid relationship target from '{}': {error}",
                    source.partname().as_str()
                ))
            })?;
            if same_name(&related, target) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn removable_images(
    package: &OpcPackage,
    removed_ribbon: &PackURI,
    candidates: &[PackURI],
) -> Result<Vec<PackURI>> {
    let names = NameIndex::new(candidates);
    let mut keep = vec![false; candidates.len()];
    let mut outgoing = vec![Vec::new(); candidates.len()];
    let mut scanned = 0usize;
    let mut edges = 0usize;

    for relationship in package.rels().iter() {
        bump_graph_links(&mut scanned)?;
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid package relationship target: {error}"))
        })?;
        if let Some(index) = names.get(&target) {
            keep[index] = true;
        }
    }

    for source in package.iter_parts() {
        if same_name(source.partname(), removed_ribbon) {
            continue;
        }
        let source_index = names.get(source.partname());
        for relationship in source.rels().iter() {
            bump_graph_links(&mut scanned)?;
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid relationship target from '{}': {error}",
                    source.partname().as_str()
                ))
            })?;
            let Some(target_index) = names.get(&target) else {
                continue;
            };
            match source_index {
                Some(source_index) => {
                    edges = edges.checked_add(1).ok_or(Error::Limit {
                        resource: "Ribbon image garbage-collection edges",
                        max: MAX_IMAGE_GC_EDGES,
                        actual: usize::MAX,
                    })?;
                    if edges > MAX_IMAGE_GC_EDGES {
                        return Err(Error::Limit {
                            resource: "Ribbon image garbage-collection edges",
                            max: MAX_IMAGE_GC_EDGES,
                            actual: edges,
                        });
                    }
                    outgoing[source_index].push(target_index);
                },
                None => keep[target_index] = true,
            }
        }
    }

    for targets in &mut outgoing {
        targets.sort_unstable();
        targets.dedup();
    }

    let mut pending: VecDeque<_> = keep
        .iter()
        .enumerate()
        .filter_map(|(index, keep)| keep.then_some(index))
        .collect();
    while let Some(source) = pending.pop_front() {
        for &target in &outgoing[source] {
            if !keep[target] {
                keep[target] = true;
                pending.push_back(target);
            }
        }
    }

    Ok(candidates
        .iter()
        .zip(keep)
        .filter(|(_, keep)| !keep)
        .map(|(candidate, _)| candidate.clone())
        .collect())
}

fn bump_graph_links(scanned: &mut usize) -> Result<()> {
    *scanned = scanned.checked_add(1).ok_or(Error::Limit {
        resource: "Ribbon package graph relationships",
        max: MAX_GRAPH_LINKS,
        actual: usize::MAX,
    })?;
    if *scanned > MAX_GRAPH_LINKS {
        return Err(Error::Limit {
            resource: "Ribbon package graph relationships",
            max: MAX_GRAPH_LINKS,
            actual: *scanned,
        });
    }
    Ok(())
}

struct NameIndex<'a> {
    values: &'a [PackURI],
    order: Vec<usize>,
}

impl<'a> NameIndex<'a> {
    fn new(values: &'a [PackURI]) -> Self {
        let mut order: Vec<_> = (0..values.len()).collect();
        order.sort_unstable_by(|left, right| compare_names(&values[*left], &values[*right]));
        Self { values, order }
    }

    fn get(&self, wanted: &PackURI) -> Option<usize> {
        self.order
            .binary_search_by(|index| compare_names(&self.values[*index], wanted))
            .ok()
            .map(|position| self.order[position])
    }
}

fn compare_names(left: &PackURI, right: &PackURI) -> Ordering {
    left.as_str()
        .bytes()
        .map(|byte| byte.to_ascii_lowercase())
        .cmp(right.as_str().bytes().map(|byte| byte.to_ascii_lowercase()))
}

fn same_name(left: &PackURI, right: &PackURI) -> bool {
    left.is_equivalent_to(right)
}
