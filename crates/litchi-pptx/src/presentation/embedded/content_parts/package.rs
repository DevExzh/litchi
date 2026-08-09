//! OPC graph ownership for inert `PresentationML` content parts.

use super::codec::scan_slide;
use super::model::{ContentPart, Payload, Relationship, RelationshipMetadata, Target};
use super::transaction::{Commit, Patch, RelationshipState, Snapshot};
use super::validation::{
    self, MAX_PAYLOAD_BYTES, MAX_PAYLOAD_RELATIONSHIPS, MAX_RELATIONSHIP_FIELD_BYTES,
    MAX_TOTAL_PAYLOAD_BYTES,
};
use crate::presentation::embedded::{invalid, limit};
use crate::{Error, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

const MAX_CONTENT_PARTS: usize = 4_096;

/// Finite resource policy for content-part discovery.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum number of anchors across calls sharing this policy.
    pub content_parts: usize,
    /// Maximum bytes retained for one unique internal payload.
    pub payload_bytes: usize,
    /// Maximum aggregate bytes retained for unique internal payloads.
    pub total_payload_bytes: usize,
    /// Maximum relationship records retained per unique payload.
    pub payload_relationships: usize,
    used_payload_bytes: usize,
}

impl Limits {
    /// Conservative bounded policy for ordinary PPTX packages.
    pub const DEFAULT: Self = Self {
        content_parts: MAX_CONTENT_PARTS,
        payload_bytes: MAX_PAYLOAD_BYTES,
        total_payload_bytes: MAX_TOTAL_PAYLOAD_BYTES,
        payload_relationships: MAX_PAYLOAD_RELATIONSHIPS,
        used_payload_bytes: 0,
    };

    /// Construct a nonzero finite policy.
    #[must_use]
    pub const fn new(
        content_parts: usize,
        payload_bytes: usize,
        total_payload_bytes: usize,
        payload_relationships: usize,
    ) -> Option<Self> {
        if content_parts == 0
            || payload_bytes == 0
            || total_payload_bytes == 0
            || payload_relationships == 0
        {
            None
        } else {
            Some(Self {
                content_parts,
                payload_bytes,
                total_payload_bytes,
                payload_relationships,
                used_payload_bytes: 0,
            })
        }
    }

    fn charge_anchor(&mut self, actual: usize) -> Result<()> {
        if actual >= self.content_parts {
            return Err(limit("content-part count", self.content_parts));
        }
        Ok(())
    }

    fn charge_payload(&mut self, bytes: usize) -> Result<()> {
        if bytes > self.payload_bytes {
            return Err(limit("content-part payload bytes", self.payload_bytes));
        }
        let actual = self
            .used_payload_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("total content-part payload bytes", self.total_payload_bytes))?;
        if actual > self.total_payload_bytes {
            return Err(limit(
                "total content-part payload bytes",
                self.total_payload_bytes,
            ));
        }
        self.used_payload_bytes = actual;
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Load every active `p:contentPart` anchor from one `PresentationML` slide.
///
/// The `r:id` on each anchor is checked against the owning slide's actual
/// relationship collection. Internal targets are copied as opaque bytes and
/// external targets are retained as URI metadata only; neither target kind is
/// interpreted, opened, rendered, or executed.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_slide(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut Limits,
) -> Result<Vec<ContentPart>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "content-part discovery requires a PresentationML slide",
        ));
    }
    let anchors = scan_slide(slide.blob(), limits.content_parts.max(1))?;
    let slide_part_name = slide.partname().clone();
    let mut payloads = HashMap::<PackURI, Payload>::new();
    let mut result = Vec::with_capacity(anchors.len());
    for (index, anchor) in anchors.into_iter().enumerate() {
        limits.charge_anchor(index)?;
        let relationship = slide.rels().get(anchor.relationship_id()).ok_or_else(|| {
            Error::Relationship(format!(
                "slide {slide_index} contentPart relationship '{}' is missing",
                anchor.relationship_id()
            ))
        })?;
        let relationship_data = relationship_metadata(relationship)?;
        let target = if relationship.is_external() {
            if relationship.target_ref().is_empty() {
                return Err(Error::Relationship(format!(
                    "slide {slide_index} contentPart relationship '{}' has an empty external target",
                    anchor.relationship_id()
                )));
            }
            Target::External {
                target_ref: relationship.target_ref().to_owned(),
            }
        } else {
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return Err(Error::Relationship(format!(
                    "slide {slide_index} contentPart relationship '{}' has an internal query or fragment",
                    anchor.relationship_id()
                )));
            }
            let part_name = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "slide {slide_index} contentPart relationship '{}' has an invalid target: {error}",
                    anchor.relationship_id()
                ))
            })?;
            let payload = if let Some(payload) = payloads.get(&part_name) {
                payload.clone()
            } else {
                let part = package.get_part(&part_name).map_err(|error| {
                    Error::PartNotFound(format!(
                        "contentPart target '{}' from slide {slide_index}: {error}",
                        part_name.as_str()
                    ))
                })?;
                limits.charge_payload(part.blob().len())?;
                let payload = make_payload(part, limits.payload_relationships)?;
                payloads.insert(part_name, payload.clone());
                payload
            };
            Target::Internal(payload)
        };
        result.push(ContentPart {
            slide_index,
            slide_part_name: slide_part_name.clone(),
            index,
            anchor,
            relationship: Relationship {
                id: relationship_data.id,
                relationship_type: relationship_data.relationship_type,
                target_ref: relationship_data.target_ref,
                target_mode: relationship_data.target_mode,
                target,
            },
        });
    }
    Ok(result)
}

/// Capture a source-checked snapshot of one slide-owned content-part graph.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_snapshot(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut Limits,
) -> Result<Snapshot> {
    let parts = load_slide(package, slide_index, slide, limits)?;
    validation::validate_package_graph(package, slide, &parts)?;
    let relationships = super::transaction::relationship_states(slide.rels().iter())?;
    Snapshot::from_parts(
        slide_index,
        slide.partname().clone(),
        slide.blob_arc(),
        parts,
        relationships,
    )
}

/// Apply a source-checked content-part patch atomically to its owning package.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let before = patch.before();
    let slide = package.get_part(&before.slide_part_name)?;
    let current = load_snapshot(package, before.slide_index, slide, &mut Limits::default())?;
    if !current.same_source(before) {
        return Err(validation::invalid_revision());
    }
    if patch.is_empty() {
        return Ok(current);
    }

    let mut staged = package.clone();
    staged.unsign();
    install_patch(&mut staged, patch)?;
    let slide = staged.get_part(&before.slide_part_name)?;
    let resulting = load_snapshot(
        &staged,
        patch.after().slide_index,
        slide,
        &mut Limits::default(),
    )?;
    if !resulting.same_source(patch.after()) {
        return Err(invalid(
            "published content-part graph differs from the commit",
        ));
    }
    *package = staged;
    Ok(resulting)
}

/// Apply a committed content-part transaction atomically.
///
/// # Errors
///
/// Returns an error if the operation fails.
#[inline]
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, commit.patch())
}

fn install_patch(package: &mut OpcPackage, patch: &Patch) -> Result<()> {
    let before = patch.before();
    let after = patch.after();
    let slide_name = before.slide_part_name.clone();
    package
        .get_part_mut(&slide_name)?
        .set_blob(after.source_xml.as_ref().clone());
    sync_slide_relationships(
        package.get_part_mut(&slide_name)?,
        &after.slide_relationships,
    )?;

    let before_payloads: HashSet<PackURI> = before
        .payloads
        .iter()
        .map(|payload| payload.part_name().clone())
        .collect();
    let after_payloads: HashSet<PackURI> = after
        .payloads
        .iter()
        .map(|payload| payload.part_name().clone())
        .collect();

    for payload in after.payloads.iter() {
        if package.contains_part(payload.part_name()) {
            if !before_payloads.contains(payload.part_name()) {
                return Err(invalid(format!(
                    "content-part payload '{}' conflicts with an unrelated package part",
                    payload.part_name().as_str()
                )));
            }
            let part = package.get_part_mut(payload.part_name())?;
            part.set_content_type(payload.content_type().to_owned())?;
            part.set_blob(payload.bytes().to_vec());
            sync_payload_relationships(part, payload.relationships())?;
        } else {
            package.try_add_part(Box::new(BlobPart::new(
                payload.part_name().clone(),
                payload.content_type().to_owned(),
                payload.bytes().to_vec(),
            )))?;
            sync_payload_relationships(
                package.get_part_mut(payload.part_name())?,
                payload.relationships(),
            )?;
        }
    }

    for part_name in before_payloads.difference(&after_payloads) {
        if !has_inbound_relationship(package, part_name)? {
            package.remove_part(part_name);
        }
    }
    Ok(())
}

fn sync_slide_relationships(part: &mut dyn Part, desired: &[RelationshipState]) -> Result<()> {
    let current: Vec<String> = part
        .rels()
        .iter()
        .map(|value| value.r_id().to_owned())
        .collect();
    for id in current {
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

fn sync_payload_relationships(part: &mut dyn Part, desired: &[RelationshipMetadata]) -> Result<()> {
    let current: Vec<String> = part
        .rels()
        .iter()
        .map(|value| value.r_id().to_owned())
        .collect();
    for id in current {
        part.rels_mut().remove(&id);
    }
    for relationship in desired {
        part.rels_mut().try_add_relationship(
            relationship.relationship_type().to_owned(),
            relationship.target_ref().to_owned(),
            relationship.id().to_owned(),
            relationship.target_mode(),
        )?;
    }
    Ok(())
}

fn has_inbound_relationship(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|value| value.is_equivalent_to(target))
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn make_payload(part: &dyn Part, maximum_relationships: usize) -> Result<Payload> {
    let mut relationships = Vec::new();
    for relationship in part.rels().iter() {
        if relationships.len() >= maximum_relationships {
            return Err(limit(
                "content-part payload relationships",
                maximum_relationships,
            ));
        }
        relationships.push(relationship_metadata(relationship)?);
    }
    relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    Ok(Payload {
        part_name: part.partname().clone(),
        content_type: Arc::<str>::from(part.content_type()),
        bytes: Arc::<[u8]>::from(part.blob()),
        relationships: Arc::<[RelationshipMetadata]>::from(relationships),
    })
}

fn relationship_metadata(relationship: &litchi_opc::Relationship) -> Result<RelationshipMetadata> {
    for (field, value) in [
        ("relationship id", relationship.r_id()),
        ("relationship type", relationship.reltype()),
        ("relationship target", relationship.target_ref()),
    ] {
        if value.len() > MAX_RELATIONSHIP_FIELD_BYTES {
            return Err(limit(
                "content-part relationship metadata bytes",
                MAX_RELATIONSHIP_FIELD_BYTES,
            ));
        }
        if field == "relationship id" && value.is_empty() {
            return Err(invalid("content-part relationship id is empty"));
        }
    }
    Ok(RelationshipMetadata {
        id: relationship.r_id().to_owned(),
        relationship_type: relationship.reltype().to_owned(),
        target_ref: relationship.target_ref().to_owned(),
        target_mode: relationship.target_mode(),
    })
}
