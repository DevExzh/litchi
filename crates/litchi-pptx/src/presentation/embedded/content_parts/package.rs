//! OPC graph ownership for inert PresentationML content parts.

use super::codec::scan_slide;
use super::model::{ContentPart, Payload, Relationship, RelationshipMetadata, Target};
use crate::presentation::embedded::{invalid, limit};
use crate::{Error, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part};
use std::collections::HashMap;
use std::sync::Arc;

const MAX_CONTENT_PARTS: usize = 4_096;
const MAX_PAYLOAD_BYTES: usize = 64 * 1024 * 1024;
const MAX_TOTAL_PAYLOAD_BYTES: usize = 256 * 1024 * 1024;
const MAX_PAYLOAD_RELATIONSHIPS: usize = 4_096;
const MAX_RELATIONSHIP_FIELD_BYTES: usize = 16 * 1024;

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

/// Load every active `p:contentPart` anchor from one PresentationML slide.
///
/// The `r:id` on each anchor is checked against the owning slide's actual
/// relationship collection. Internal targets are copied as opaque bytes and
/// external targets are retained as URI metadata only; neither target kind is
/// interpreted, opened, rendered, or executed.
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
