//! Package-facing snapshot relationship and content-type codec helpers.

use super::super::model::{AddIn, Limits, SnapshotResource, SnapshotTarget};
use super::super::package::{PackageGraphIndex, fold_part_name};
use super::super::validation::validate_image_content_type;
use super::super::{IMAGE_RELATIONSHIP_TYPE, STRICT_IMAGE_RELATIONSHIP_TYPE};
use super::semantic::{invalid, limit};
use crate::{Error, Result};
use litchi_opc::{OpcPackage, PackURI, Part};
use std::collections::{HashMap, HashSet};

pub(in crate::web) fn load_snapshot_resources(
    package: &OpcPackage,
    part: &dyn Part,
    extension: &AddIn,
    total_snapshot_bytes: &mut usize,
    counted_snapshot_parts: &mut HashSet<String>,
    limits: &Limits,
    index: &PackageGraphIndex,
) -> Result<Vec<SnapshotResource>> {
    let mut referenced = HashMap::new();
    if let Some(snapshot) = &extension.snapshot {
        if let Some(id) = &snapshot.embedded_relationship_id
            && referenced.insert(id.as_str(), false).is_some()
        {
            return invalid("snapshot embed and link IDs must differ".into());
        }
        if let Some(id) = &snapshot.linked_relationship_id
            && referenced.insert(id.as_str(), true).is_some()
        {
            return invalid("snapshot embed and link IDs must differ".into());
        }
    }
    let mut resources = Vec::with_capacity(referenced.len());
    for relationship in part.rels().iter() {
        let Some(linked) = referenced.remove(relationship.r_id()) else {
            return invalid(format!(
                "web extension part has unreferenced relationship '{}'",
                relationship.r_id()
            ));
        };
        if !matches!(
            relationship.reltype(),
            IMAGE_RELATIONSHIP_TYPE | STRICT_IMAGE_RELATIONSHIP_TYPE
        ) {
            return invalid(format!(
                "snapshot relationship '{}' is not an image relationship",
                relationship.r_id()
            ));
        }
        if relationship.is_external() {
            if !linked {
                return invalid(format!(
                    "embedded snapshot relationship '{}' must be internal",
                    relationship.r_id()
                ));
            }
            resources.push(SnapshotResource {
                relationship_id: relationship.r_id().to_owned(),
                target: SnapshotTarget::External {
                    target: relationship.target_ref().to_owned(),
                },
            });
            continue;
        }
        let image_target = checked_internal_target(relationship, "snapshot image")?;
        let image_name = index
            .canonical(&image_target)
            .ok_or_else(|| Error::Missing(format!("snapshot image '{}'", image_target.as_str())))?;
        let image = package.get_part(image_name).map_err(|error| {
            Error::Missing(format!("snapshot image '{}': {error}", image_name.as_str()))
        })?;
        validate_image_content_type(image.content_type())?;
        if image.rels().iter().next().is_some() {
            return invalid(format!(
                "snapshot image '{}' must not have relationships",
                image_name.as_str()
            ));
        }
        if image.blob().len() > limits.image_bytes {
            return limit(
                "web extension snapshot bytes",
                limits.image_bytes,
                image.blob().len(),
            );
        }
        let image_name = image.partname().clone();
        if counted_snapshot_parts.insert(fold_part_name(&image_name)) {
            *total_snapshot_bytes = total_snapshot_bytes
                .checked_add(image.blob().len())
                .ok_or_else(|| Error::Invalid("aggregate snapshot byte count overflow".into()))?;
            if *total_snapshot_bytes > limits.total_image_bytes {
                return limit(
                    "aggregate web extension snapshot bytes",
                    limits.total_image_bytes,
                    *total_snapshot_bytes,
                );
            }
        }
        resources.push(SnapshotResource {
            relationship_id: relationship.r_id().to_owned(),
            target: SnapshotTarget::Internal {
                part_name: image_name,
                content_type: image.content_type().to_owned(),
                data: image.blob_arc(),
            },
        });
    }
    if let Some((id, _)) = referenced.into_iter().next() {
        return invalid(format!("snapshot references missing relationship '{id}'"));
    }
    let embedded_id = extension
        .snapshot
        .as_ref()
        .and_then(|snapshot| snapshot.embedded_relationship_id.as_deref());
    resources.sort_by(|left, right| {
        let left_order = usize::from(Some(left.relationship_id.as_str()) != embedded_id);
        let right_order = usize::from(Some(right.relationship_id.as_str()) != embedded_id);
        left_order
            .cmp(&right_order)
            .then_with(|| left.relationship_id.cmp(&right.relationship_id))
    });
    Ok(resources)
}

pub(in crate::web) fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() != expected {
        Err(Error::ContentType {
            expected: expected.into(),
            actual: part.content_type().into(),
        })
    } else {
        Ok(())
    }
}

pub(in crate::web) fn checked_internal_target(
    relationship: &litchi_opc::Relationship,
    label: &str,
) -> Result<PackURI> {
    if relationship.is_external() {
        return Err(Error::Relationship(format!(
            "{label} relationship '{}' must be internal",
            relationship.r_id()
        )));
    }
    if relationship.target_ref().contains(['?', '#']) {
        return Err(Error::Relationship(format!(
            "{label} relationship '{}' has an internal target with a query or fragment",
            relationship.r_id()
        )));
    }
    relationship.target_partname().map_err(|error| {
        Error::Relationship(format!(
            "invalid {label} relationship target '{}': {error}",
            relationship.r_id()
        ))
    })
}
