//! Exact source fingerprints shared by typed XLSB package transactions.

use std::hash::{Hash, Hasher};
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use super::error::Result;

const MAX_TRANSACTION_PARTS: usize = 65_536;
const MAX_TRANSACTION_RELATIONSHIPS: usize = 1_000_000;
const MAX_TRANSACTION_STRING_BYTES: usize = 1_048_576;

/// Borrowed ASCII-case-insensitive key without a lowercase allocation.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Caseless<'a>(&'a str);

impl<'a> Caseless<'a> {
    pub(crate) const fn new(value: &'a str) -> Self {
        Self(value)
    }
}

impl PartialEq for Caseless<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.0.eq_ignore_ascii_case(other.0)
    }
}

impl Eq for Caseless<'_> {}

impl Hash for Caseless<'_> {
    fn hash<H: Hasher>(&self, state: &mut H) {
        for byte in self.0.bytes() {
            state.write_u8(byte.to_ascii_lowercase());
        }
    }
}

/// Exact bytes and relationship topology for a typed owner dependency closure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Source {
    root_relationships: Vec<SourceRelationship>,
    parts: Vec<SourcePart>,
}

impl Source {
    pub(crate) fn capture(package: &OpcPackage) -> Result<Self> {
        let root_relationships = capture_relationships(package.rels())?;
        let part_count = package.iter_parts().count();
        if part_count > MAX_TRANSACTION_PARTS {
            return Err(super::error::Error::InvalidLength {
                expected: MAX_TRANSACTION_PARTS,
                found: part_count,
            });
        }
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(part_count)
            .map_err(|source| super::error::Error::Allocation {
                resource: "XLSB transaction source parts",
                source,
            })?;
        for part in package.iter_parts() {
            parts.push(SourcePart::capture(part)?);
        }
        parts.sort_unstable_by(|left, right| left.name.cmp(&right.name));
        Ok(Self {
            root_relationships,
            parts,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SourcePart {
    name: String,
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<SourceRelationship>,
}

impl SourcePart {
    fn capture(part: &dyn Part) -> Result<Self> {
        Ok(Self {
            name: try_string(part.partname().as_str(), "XLSB transaction part name")?,
            content_type: try_string(part.content_type(), "XLSB transaction content type")?,
            blob: part.blob_arc(),
            relationships: capture_relationships(part.rels())?,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SourceRelationship {
    id: String,
    relationship_type: String,
    target: String,
    mode: TargetMode,
}

fn capture_relationships(
    relationships: &litchi_opc::Relationships,
) -> Result<Vec<SourceRelationship>> {
    let count = relationships.iter().count();
    if count > MAX_TRANSACTION_RELATIONSHIPS {
        return Err(super::error::Error::InvalidLength {
            expected: MAX_TRANSACTION_RELATIONSHIPS,
            found: count,
        });
    }
    let mut captured = Vec::new();
    captured
        .try_reserve_exact(count)
        .map_err(|source| super::error::Error::Allocation {
            resource: "XLSB transaction relationships",
            source,
        })?;
    for relationship in relationships.iter() {
        captured.push(SourceRelationship {
            id: try_string(relationship.r_id(), "XLSB transaction relationship id")?,
            relationship_type: try_string(
                relationship.reltype(),
                "XLSB transaction relationship type",
            )?,
            target: try_string(
                relationship.target_ref(),
                "XLSB transaction relationship target",
            )?,
            mode: relationship.target_mode(),
        });
    }
    captured.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    Ok(captured)
}

fn try_string(value: &str, resource: &'static str) -> Result<String> {
    if value.len() > MAX_TRANSACTION_STRING_BYTES {
        return Err(super::error::Error::InvalidLength {
            expected: MAX_TRANSACTION_STRING_BYTES,
            found: value.len(),
        });
    }
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| super::error::Error::Allocation { resource, source })?;
    output.push_str(value);
    Ok(output)
}

/// Refuse deletion or replacement when an owned target has another inbound edge.
pub(crate) fn require_exclusive_inbound(
    package: &OpcPackage,
    owner_part: &PackURI,
    targets: &[(String, PackURI)],
    description: &'static str,
) -> Result<()> {
    for relationship in package.rels().iter() {
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname()?;
        if targets
            .iter()
            .any(|(_, owned_target)| *owned_target == target)
        {
            return Err(super::error::Error::InvalidRelationship(format!(
                "{description} target {target} is also referenced from the package root"
            )));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            let Some((relationship_id, _)) = targets
                .iter()
                .find(|(_, owned_target)| *owned_target == target)
            else {
                continue;
            };
            if source.partname() != owner_part || relationship.r_id() != relationship_id {
                return Err(super::error::Error::InvalidRelationship(format!(
                    "{description} target {target} has a non-owner inbound relationship"
                )));
            }
        }
    }
    Ok(())
}
