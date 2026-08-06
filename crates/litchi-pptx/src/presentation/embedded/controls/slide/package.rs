//! OPC ownership and atomic publication for one slide-owned ActiveX graph.

use std::collections::HashSet;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationship};

use super::super::model::{Binary, Control};
use super::super::{
    BINARY_CONTENT_TYPE, BINARY_RELATIONSHIP, CONTROL_RELATIONSHIP, DESCRIPTOR_CONTENT_TYPE,
    Limits, STRICT_CONTROL_RELATIONSHIP,
};
use super::transaction::{BinarySource, Commit, Patch, RelationshipState, Snapshot};
use super::validation::invalid_revision;
use crate::presentation::embedded::invalid;
use crate::{Error, Result};

/// Load one typed control snapshot from its owning slide.
pub fn load(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    control_index: usize,
    limits: &mut Limits,
) -> Result<Snapshot> {
    let controls = super::super::load_slide(package, slide_index, slide, limits)?;
    validate_control_relationships(slide, &controls)?;
    let control = controls
        .get(control_index)
        .cloned()
        .ok_or_else(|| Error::IndexOutOfBounds {
            index: control_index,
            len: controls.len(),
        })?;
    let slide_part_name = slide.partname().clone();
    let slide_relationships = relationship_states(slide.rels().iter())?;

    let (descriptor_xml, descriptor_relationships, binary) = match control.descriptor.as_ref() {
        None => (None, None, None),
        Some(descriptor) => {
            let descriptor_part = package.get_part(descriptor.part_name())?;
            if descriptor_part.content_type() != DESCRIPTOR_CONTENT_TYPE {
                return Err(Error::ContentType {
                    expected: DESCRIPTOR_CONTENT_TYPE.to_owned(),
                    actual: descriptor_part.content_type().to_owned(),
                });
            }
            validate_internal_targets(package, descriptor_part)?;
            let descriptor_relationships = relationship_states(descriptor_part.rels().iter())?;
            let binary = descriptor
                .binary()
                .map(|value| load_binary(package, descriptor_part, value))
                .transpose()?;
            (
                Some(descriptor_part.blob_arc()),
                Some(descriptor_relationships),
                binary,
            )
        },
    };

    Snapshot::from_parts(
        slide_index,
        control_index,
        slide_part_name,
        control,
        slide.blob_arc(),
        slide_relationships,
        descriptor_xml,
        descriptor_relationships,
        binary,
    )
}

/// Apply a source-checked patch to its owning slide atomically.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let before = patch.before();
    let slide = package.get_part(&before.slide_part_name)?;
    let current = load(
        package,
        before.slide_index,
        slide,
        before.control_index,
        &mut Limits::default(),
    )?;
    if !current.same_source(before) {
        return Err(invalid_revision());
    }
    if patch.is_empty() {
        return Ok(current);
    }

    let mut staged = package.clone();
    staged.unsign();
    install_patch(&mut staged, patch)?;
    let slide = staged.get_part(&before.slide_part_name)?;
    let resulting = load(
        &staged,
        patch.after().slide_index,
        slide,
        patch.after().control_index,
        &mut Limits::default(),
    )?;
    if !resulting.same_source(patch.after()) {
        return Err(invalid(
            "ActiveX patch target failed slide relationship or descriptor validation",
        ));
    }
    *package = staged;
    Ok(resulting)
}

/// Apply a committed transaction and return the validated post-publication snapshot.
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    let patch = commit.into_patch();
    apply_patch(package, &patch)
}

fn install_patch(package: &mut OpcPackage, patch: &Patch) -> Result<()> {
    let before = patch.before();
    let after = patch.after();
    let slide = package.get_part_mut(&before.slide_part_name)?;
    slide.set_blob(after.source_xml.as_ref().clone());

    let before_descriptor = before.control.descriptor.as_ref();
    let after_descriptor = after.control.descriptor.as_ref();
    if before_descriptor.map(|value| value.part_name())
        != after_descriptor.map(|value| value.part_name())
    {
        return Err(invalid(
            "ActiveX descriptor identity cannot change in a patch",
        ));
    }
    let Some(descriptor) = after_descriptor else {
        if before_descriptor.is_some() || after.descriptor_xml.is_some() {
            return Err(invalid("ActiveX descriptor graph is incomplete"));
        }
        return Ok(());
    };
    let descriptor_name = descriptor.part_name().clone();
    if let Some(binary) = after.binary.as_ref() {
        ensure_binary_part(package, binary, before.binary.as_ref())?;
    }
    let descriptor_xml = after
        .descriptor_xml
        .as_ref()
        .ok_or_else(|| invalid("ActiveX descriptor XML is missing"))?;
    {
        let descriptor_part = package.get_part_mut(&descriptor_name)?;
        if let Some(binary) = after.binary.as_ref() {
            ensure_relationship(descriptor_part, binary)?;
        } else if let Some(binary) = before.binary.as_ref() {
            descriptor_part.rels_mut().remove(&binary.relationship_id);
        }
        let actual_relationships = relationship_states(descriptor_part.rels().iter())?;
        if Some(&actual_relationships) != after.descriptor_relationships.as_deref() {
            return Err(invalid(
                "ActiveX descriptor relationship lifecycle does not match the patch target",
            ));
        }
        descriptor_part.set_blob(descriptor_xml.as_ref().clone());
    }

    if after.binary.is_none() {
        if let Some(binary) = before.binary.as_ref() {
            if !has_inbound_relationship(package, &binary.part_name)? {
                package.remove_part(&binary.part_name);
            }
        }
    }
    Ok(())
}

fn load_binary(
    package: &OpcPackage,
    descriptor: &dyn Part,
    binary: &Binary,
) -> Result<BinarySource> {
    let relationship = descriptor
        .rels()
        .get(binary.relationship_id())
        .ok_or_else(|| Error::Relationship("ActiveX binary relationship is missing".into()))?;
    if relationship.is_external() || relationship.reltype() != BINARY_RELATIONSHIP {
        return Err(Error::Relationship(
            "ActiveX binary relationship has an unsupported type".into(),
        ));
    }
    let part_name = relationship.target_partname()?;
    if part_name != *binary.part_name() {
        return Err(Error::Relationship(
            "ActiveX binary relationship target does not match its descriptor".into(),
        ));
    }
    let part = package.get_part(&part_name)?;
    if part.content_type() != BINARY_CONTENT_TYPE {
        return Err(Error::ContentType {
            expected: BINARY_CONTENT_TYPE.to_owned(),
            actual: part.content_type().to_owned(),
        });
    }
    Ok(BinarySource {
        relationship_id: relationship.r_id().to_owned(),
        relationship_type: relationship.reltype().to_owned(),
        target_ref: relationship.target_ref().to_owned(),
        target_mode: relationship.target_mode(),
        part_name,
        content_type: part.content_type().to_owned(),
        bytes: part.blob_arc(),
        relationships: Arc::new(relationship_states(part.rels().iter())?),
    })
}

fn validate_control_relationships(slide: &dyn Part, controls: &[Control]) -> Result<()> {
    let referenced: HashSet<&str> = controls
        .iter()
        .filter_map(|control| control.relationship_id())
        .collect();
    for relationship in slide.rels().iter() {
        if matches!(
            relationship.reltype(),
            CONTROL_RELATIONSHIP | STRICT_CONTROL_RELATIONSHIP
        ) && !referenced.contains(relationship.r_id())
        {
            return Err(Error::Relationship(format!(
                "slide contains unreferenced ActiveX control relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    Ok(())
}

fn validate_internal_targets(package: &OpcPackage, part: &dyn Part) -> Result<()> {
    for relationship in part.rels().iter() {
        if relationship.is_external() {
            if relationship.target_ref().is_empty() {
                return Err(Error::Relationship(format!(
                    "external ActiveX relationship '{}' has an empty target",
                    relationship.r_id()
                )));
            }
        } else {
            package.get_part(&relationship.target_partname()?)?;
        }
    }
    Ok(())
}

fn relationship_states<'a>(
    relationships: impl Iterator<Item = &'a Relationship>,
) -> Result<Vec<RelationshipState>> {
    relationships
        .map(|relationship| {
            Ok(RelationshipState {
                id: relationship.r_id().to_owned(),
                relationship_type: relationship.reltype().to_owned(),
                target_ref: relationship.target_ref().to_owned(),
                target_mode: relationship.target_mode(),
            })
        })
        .collect()
}

fn ensure_relationship(part: &mut dyn Part, desired: &BinarySource) -> Result<()> {
    if let Some(current) = part.rels().get(&desired.relationship_id) {
        if current.reltype() != desired.relationship_type
            || current.target_ref() != desired.target_ref
            || current.target_mode() != desired.target_mode
        {
            return Err(Error::Relationship(format!(
                "ActiveX binary relationship '{}' conflicts with the patch target",
                desired.relationship_id
            )));
        }
        return Ok(());
    }
    part.rels_mut().try_add_relationship(
        desired.relationship_type.clone(),
        desired.target_ref.clone(),
        desired.relationship_id.clone(),
        desired.target_mode,
    )?;
    Ok(())
}

fn ensure_binary_part(
    package: &mut OpcPackage,
    desired: &BinarySource,
    before: Option<&BinarySource>,
) -> Result<()> {
    if package.contains_part(&desired.part_name) {
        let shared_update = before
            .is_some_and(|value| value.bytes.as_ref() != desired.bytes.as_ref())
            && has_multiple_inbound_relationships(package, &desired.part_name)?;
        let part = package.get_part_mut(&desired.part_name)?;
        if part.content_type() != desired.content_type {
            return Err(Error::ContentType {
                expected: desired.content_type.clone(),
                actual: part.content_type().to_owned(),
            });
        }
        let current_relationships = relationship_states(part.rels().iter())?;
        if current_relationships != desired.relationships.as_ref().clone() {
            return Err(invalid("ActiveX binary part relationships are stale"));
        }
        if shared_update {
            return Err(invalid(
                "cannot replace a shared ActiveX binary payload through one control",
            ));
        }
        part.set_blob(desired.bytes.as_ref().clone());
        return Ok(());
    }
    package.validate_new_part_name(&desired.part_name)?;
    let mut part = BlobPart::new(
        desired.part_name.clone(),
        desired.content_type.clone(),
        desired.bytes.as_ref().clone(),
    );
    for relationship in desired.relationships.iter() {
        part.rels_mut().try_add_relationship(
            relationship.relationship_type.clone(),
            relationship.target_ref.clone(),
            relationship.id.clone(),
            relationship.target_mode,
        )?;
    }
    package.try_add_part(Box::new(part))?;
    Ok(())
}

fn has_inbound_relationship(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    Ok(inbound_relationship_count(package, target)? != 0)
}

fn has_multiple_inbound_relationships(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    Ok(inbound_relationship_count(package, target)? > 1)
}

fn inbound_relationship_count(package: &OpcPackage, target: &PackURI) -> Result<usize> {
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
