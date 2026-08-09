use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part, Relationship};

use super::super::Limits;
use super::super::package::load_slide;
use super::transaction::{
    Commit, PartSource, Patch, RelationshipState, Snapshot, sorted_parts, sorted_relationships,
};
use super::validation::{MAX_RELATIONSHIPS, validate_source};
use crate::presentation::embedded::invalid;
use crate::{Error, Result};

/// Capture one slide-owned OLE graph after validating ownership and targets.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(
    package: &OpcPackage,
    slide_index: usize,
    slide_part_name: &PackURI,
) -> Result<Snapshot> {
    let slide = package.get_part(slide_part_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_owned(),
            actual: slide.content_type().to_owned(),
        });
    }
    validate_source(slide.blob())?;
    let mut limits = Limits::default();
    let objects = load_slide(package, slide_index, slide, &mut limits)?;
    validate_slide_relationships(slide, &objects)?;
    let parts = capture_parts(package, slide, &objects)?;
    for part in package.iter_parts() {
        if matches!(part.content_type(), ct::OFC_OLE_OBJECT | ct::OFC_PACKAGE)
            && inbound_count(package, part.partname())? == 0
        {
            return Err(Error::Relationship(format!(
                "OLE payload part '{}' is orphaned",
                part.partname()
            )));
        }
    }
    let package_part_names = package
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    Snapshot::from_parts(
        slide_index,
        slide.partname().clone(),
        slide.blob_arc(),
        relationship_states(slide.rels().iter())?,
        objects,
        parts,
        package_part_names,
    )
}

/// Publish a committed OLE graph replacement atomically.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let current = load(
        package,
        patch.before().slide_index,
        &patch.before().slide_part_name,
    )?;
    if !current.same_source(patch.before()) {
        return Err(super::validation::invalid_revision());
    }
    if patch.is_empty() {
        return Ok(current);
    }
    let mut candidate = package.clone();
    candidate.unsign();
    install(&mut candidate, patch)?;
    let resulting = load(
        &candidate,
        patch.after().slide_index,
        &patch.after().slide_part_name,
    )?;
    if !resulting.same_source(patch.after()) {
        return Err(invalid(
            "published OLE graph differs from the commit target",
        ));
    }
    *package = candidate;
    Ok(resulting)
}

/// # Errors
///
/// Returns an error if the operation fails.
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, &commit.into_patch())
}

fn install(package: &mut OpcPackage, patch: &Patch) -> Result<()> {
    let before = patch.before();
    let after = patch.after();
    if before.slide_part_name != after.slide_part_name || before.slide_index != after.slide_index {
        return Err(invalid("OLE patch changes its slide owner"));
    }
    {
        let slide = package.get_part_mut(&before.slide_part_name)?;
        slide.set_blob_shared(Arc::clone(&after.source_xml));
        restore_relationships(slide, after.relationships.as_ref())?;
    }

    let before_parts: HashMap<PackURI, PartSource> = before
        .parts
        .iter()
        .cloned()
        .map(|part| (part.part_name.clone(), part))
        .collect();
    for desired in after.parts.iter() {
        if package.contains_part(&desired.part_name) {
            let (actual_content_type, actual_relationships, changed) = {
                let part = package.get_part(&desired.part_name)?;
                (
                    part.content_type().to_owned(),
                    relationship_states(part.rels().iter())?,
                    part.blob() != desired.bytes.as_slice(),
                )
            };
            if actual_content_type != desired.content_type {
                return Err(Error::ContentType {
                    expected: desired.content_type.clone(),
                    actual: actual_content_type,
                });
            }
            if actual_relationships != desired.relationships.as_ref().clone() {
                return Err(invalid("OLE payload part relationships are stale"));
            }
            if changed && inbound_count(package, &desired.part_name)? > 1 {
                return Err(invalid(
                    "cannot replace a shared OLE payload through one slide",
                ));
            }
            let part = package.get_part_mut(&desired.part_name)?;
            part.set_blob_shared(Arc::clone(&desired.bytes));
        } else {
            package.validate_new_part_name(&desired.part_name)?;
            let mut part = BlobPart::new(
                desired.part_name.clone(),
                desired.content_type.clone(),
                desired.bytes.as_ref().clone(),
            );
            restore_relationships(&mut part, desired.relationships.as_ref())?;
            package.try_add_part(Box::new(part))?;
        }
    }

    let after_names: HashSet<PackURI> = after
        .parts
        .iter()
        .map(|part| part.part_name.clone())
        .collect();
    for part_name in before_parts.keys() {
        if !after_names.contains(part_name)
            && !has_inbound_relationship(package, part_name)?
            && package.remove_part(part_name)
        {
            continue;
        }
    }
    Ok(())
}

fn capture_parts(
    package: &OpcPackage,
    slide: &dyn Part,
    objects: &[super::super::model::Object],
) -> Result<Vec<PartSource>> {
    let mut names = HashSet::new();
    for object in objects {
        if let Some(super::super::model::Target::Internal { part_name, .. }) = object.target() {
            names.insert(part_name.clone());
        }
        if let Some(id) = object.preview_relationship_id() {
            let relationship = slide.rels().get(id).ok_or_else(|| {
                Error::Relationship(format!("OLE preview relationship '{id}' is missing"))
            })?;
            if !relationship.is_external() {
                names.insert(relationship.target_partname()?);
            }
        }
    }
    let mut parts = Vec::with_capacity(names.len());
    for part_name in names {
        let part = package.get_part(&part_name)?;
        parts.push(PartSource {
            part_name: part.partname().clone(),
            content_type: part.content_type().to_owned(),
            bytes: part.blob_arc(),
            relationships: Arc::new(relationship_states(part.rels().iter())?),
        });
    }
    sorted_parts(parts)
}

fn validate_slide_relationships(
    slide: &dyn Part,
    objects: &[super::super::model::Object],
) -> Result<()> {
    let referenced: HashSet<&str> = objects
        .iter()
        .filter_map(|object| object.relationship_id())
        .chain(
            objects
                .iter()
                .filter_map(|object| object.preview_relationship_id()),
        )
        .collect();
    for relationship in slide.rels().iter() {
        if is_ole_relationship(relationship.reltype()) && !referenced.contains(relationship.r_id())
        {
            return Err(Error::Relationship(format!(
                "slide contains unreferenced OLE relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    Ok(())
}

fn relationship_states<'a>(
    relationships: impl Iterator<Item = &'a Relationship>,
) -> Result<Vec<RelationshipState>> {
    if relationships
        .size_hint()
        .1
        .is_some_and(|value| value > MAX_RELATIONSHIPS)
    {
        return Err(crate::presentation::embedded::limit(
            "OLE relationship count",
            MAX_RELATIONSHIPS,
        ));
    }
    sorted_relationships(
        relationships
            .map(|relationship| RelationshipState {
                id: relationship.r_id().to_owned(),
                relationship_type: relationship.reltype().to_owned(),
                target_ref: relationship.target_ref().to_owned(),
                target_mode: relationship.target_mode(),
            })
            .collect(),
    )
}

fn restore_relationships(part: &mut dyn Part, desired: &[RelationshipState]) -> Result<()> {
    let ids: Vec<String> = part
        .rels()
        .iter()
        .map(|value| value.r_id().to_owned())
        .collect();
    for id in ids {
        part.rels_mut().remove(&id);
    }
    for relationship in desired {
        part.rels_mut().try_add_relationship(
            relationship.relationship_type.clone(),
            relationship.target_ref.clone(),
            relationship.id.clone(),
            relationship.target_mode,
        )?;
    }
    Ok(())
}

fn is_ole_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::OLE_OBJECT | rt::PACKAGE | rt::STRICT_OLE_OBJECT | rt::STRICT_PACKAGE
    )
}

fn inbound_count(package: &OpcPackage, target: &PackURI) -> Result<usize> {
    let mut count = 0usize;
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()? == *target {
            count += 1;
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external() && relationship.target_partname()? == *target {
                count += 1;
            }
        }
    }
    Ok(count)
}

fn has_inbound_relationship(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    Ok(inbound_count(package, target)? != 0)
}
